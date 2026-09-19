import os
import re
import sqlite3
from threading import Thread
from flask import Flask
import pandas as pd
import telebot
from telebot import types

# ==================== خادم وهمي لمنصة Render (Web Service Port) ====================
app = Flask('')


@app.route('/')
def home():
  return 'Bot is running smoothly on Render Web Service!'


def run():
  # Render يمرر رقم المنفذ تلقائياً عبر متغير البيئة PORT
  port = int(os.environ.get('PORT', 8080))
  app.run(host='0.0.0.0', port=port)


def keep_alive():
  t = Thread(target=run)
  t.start()


# تشغيل خادم الويب أولاً لإرضاء Render
keep_alive()

# ==================== الإعدادات الأساسية ====================
BOT_TOKEN = os.environ.get(
    'BOT_TOKEN', 'ضع_التوكن_هنا_إن_لم_تستخدم_متغيرات_البيئة'
)
ADMIN_ID = int(os.environ.get('ADMIN_ID', '123456789'))

bot = telebot.TeleBot(BOT_TOKEN)

# معرفات المجموعات والقنوات لكل سنة (Chat IDs)
YEAR_CHATS = {
    "السنة الأولى": -1003953300954,
    "السنة الثانية": -1004391353751,
    "السنة الثالثة": -1004495552551,
    "السنة الرابعة": -1004364259833,
    "السنة الخامسة": -1004363731175
}


# ==================== تهيئة قاعدة البيانات ====================
def init_db():
  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            student_id TEXT PRIMARY KEY,
            name TEXT,
            year TEXT,
            phone TEXT UNIQUE,
            telegram_user_id INTEGER UNIQUE,
            joined INTEGER DEFAULT 0,
            joined_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
  conn.commit()
  conn.close()


init_db()


# ==================== دالة تنظيف أرقام الهواتف ====================
def clean_phone(phone_str):
  if not phone_str:
    return ''
  digits = re.sub(r'\D', '', str(phone_str))
  if digits.startswith('963'):
    digits = '0' + digits[3:]
  elif digits.startswith('00963'):
    digits = '0' + digits[5:]
  if len(digits) == 9 and digits.startswith('9'):
    digits = '0' + digits
  return digits


# ==================== استقبال ملف الـ Excel من الأدمن ====================
@bot.message_handler(content_types=['document'])
def handle_excel_upload(message):
  if message.from_user.id != ADMIN_ID:
    bot.reply_to(message, '⚠️ هذه الخاصية مخصصة لرئيس الهيئة/الأدمن فقط.')
    return

  file_name = message.document.file_name
  if not (file_name.endswith('.xlsx') or file_name.endswith('.xls')):
    bot.reply_to(message, '❌ يرجى إرسال ملف Excel بصيغة xlsx أو xls.')
    return

  msg = bot.reply_to(
      message,
      '⏳ جاري معالجة الملف، تنظيف البيانات، وتحديث قاعدة البيانات...',
  )

  try:
    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    excel_path = 'uploaded_responses.xlsx'
    with open(excel_path, 'wb') as new_file:
      new_file.write(downloaded_file)

    # قراءة الملف بواسطة Pandas
    df = pd.read_excel(excel_path)

    # الاعتماد على ترتيب الأعمدة: [Timestamp, Name, Student_ID, Year, Phone]
    df = df.iloc[:, [1, 2, 3, 4]]
    df.columns = ['Name', 'Student_ID', 'Year', 'Phone']

    # معالجة النصوص وتوحيد الأرقام
    df['Student_ID'] = df['Student_ID'].astype(str).str.strip()
    df['Phone'] = df['Phone'].apply(clean_phone)
    df['Year'] = df['Year'].astype(str).str.strip()

    # إزالة التكرارات بناءً على الرقم الجامعي (الاحتفاظ بأحدث استجابة)
    df_clean = df.drop_duplicates(subset=['Student_ID'], keep='last')

    conn = sqlite3.connect('telecom_students.db')
    cursor = conn.cursor()

    added_count = 0

    for _, row in df_clean.iterrows():
      try:
        cursor.execute(
            """
                    INSERT INTO students (student_id, name, year, phone)
                    VALUES (?, ?, ?, ?)
                    ON CONFLICT(student_id) DO UPDATE SET
                        name=excluded.name,
                        year=excluded.year,
                        phone=excluded.phone
                """,
            (row['Student_ID'], row['Name'], row['Year'], row['Phone']),
        )
        added_count += 1
      except sqlite3.IntegrityError:
        continue

    conn.commit()
    conn.close()

    bot.edit_message_text(
        f'✅ **تم تحديث قاعدة البيانات بنجاح!**\n\n'
        f'📊 **عدد الطلاب المجهزين في النظام:** {added_count}\n'
        f'🚀 البوت جاهز الآن للتحقق والتوزيع.',
        chat_id=message.chat.id,
        message_id=msg.message_id,
        parse_mode='Markdown',
    )

  except Exception as e:
    bot.edit_message_text(
        f'❌ حدث خطأ أثناء معالجة الملف: {str(e)}',
        chat_id=message.chat.id,
        message_id=msg.message_id,
    )


# ==================== معالجة أمر /start والتفاعل مع الطلاب ====================
@bot.message_handler(commands=['start'])
def start_command(message):
  user_id = message.from_user.id

  # 🔒 حماية رقم 1: الفحص المباشر بـ Telegram User ID
  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()
  cursor.execute(
      'SELECT name, year, joined FROM students WHERE telegram_user_id = ?',
      (user_id,),
  )
  existing_user_by_id = cursor.fetchone()

  if existing_user_by_id and existing_user_by_id[2] == 1:
    bot.send_message(
        message.chat.id,
        f'🚫 **عذراً يا {existing_user_by_id[0]}!**\n\n'
        f'لقد حصلت على رابط الانضمام لسنتك الدراسية (**{existing_user_by_id[1]}**) سابقاً.\n'
        f'⚠️ **النظام يمنع الحصول على رابط آخر أو الانضمام لسنة أخرى.**',
        parse_mode='Markdown',
    )
    conn.close()
    return

  conn.close()

  # إنشاء زر مشاركة جهة الاتصال
  markup = types.ReplyKeyboardMarkup(
      row_width=1, resize_keyboard=True, one_time_keyboard=True
  )
  button = types.KeyboardButton(
      text='📱 مشاركة رقم التليغرام لتأكيد الهوية', request_contact=True
  )
  markup.add(button)

  welcome_text = (
      'أهلاً بك في البوت الرسمي لقسم الهندسة الإلكترونية والاتصالات (الهمك)'
      ' 🎓\n\nللحصول على رابط المجموعة/القناة الخاصة بسنتك الدراسية، يرجى'
      ' الضغط على الزر أدناه لمشاركة رقمك المعتمد في الاستبيان.'
  )
  bot.send_message(message.chat.id, welcome_text, reply_markup=markup)


# ==================== استقبال رقم الهاتف والتحقق ====================
@bot.message_handler(content_types=['contact'])
def handle_contact(message):
  if not message.contact:
    return

  raw_phone = message.contact.phone_number
  cleaned_phone = clean_phone(raw_phone)
  user_id = message.from_user.id

  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()

  # 🔒 حماية رقم 2: الفحص بـ Telegram User ID
  cursor.execute(
      'SELECT name, year, joined FROM students WHERE telegram_user_id = ?',
      (user_id,),
  )
  already_joined_account = cursor.fetchone()
  if already_joined_account and already_joined_account[2] == 1:
    bot.send_message(
        message.chat.id,
        '🚫 **عذراً! حسابك التليغرام مسجّل ومستلم لرابط الانضمام سابقاً.**',
        reply_markup=types.ReplyKeyboardRemove(),
        parse_mode='Markdown',
    )
    conn.close()
    return

  # البحث عن الرقم المرفق بالاستبيان
  cursor.execute(
      'SELECT student_id, name, year, joined FROM students WHERE phone = ?',
      (cleaned_phone,),
  )
  student = cursor.fetchone()

  if not student:
    bot.send_message(
        message.chat.id,
        '❌ **لم يتم العثور على هذا الرقم في قائمة الاستبيان.**\n\n'
        'يرجى التأكد من تعبئة الاستبيان بنفس هذا الرقم، أو التواصل مع الهيئة'
        ' لمراجعة بياناتك.',
        reply_markup=types.ReplyKeyboardRemove(),
        parse_mode='Markdown',
    )
    conn.close()
    return

  student_id, name, year_name, joined = student

  # 🔒 حماية رقم 3: التأكد أن هذا الرقم الجامعي لم يُستغَل من قبل شخص آخر
  if joined == 1:
    bot.send_message(
        message.chat.id,
        f'⚠️ **عذراً، هذا الرقم الجامعي ({student_id}) استلم رابط الانضمام'
        ' سابقاً ولا يمكن استخدامه مجدداً.**',
        reply_markup=types.ReplyKeyboardRemove(),
        parse_mode='Markdown',
    )
    conn.close()
    return

  chat_target_id = YEAR_CHATS.get(year_name)

  if not chat_target_id:
    bot.send_message(
        message.chat.id,
        '⚠️ خطأ في إعدادات السنة الدراسية، يرجى مراجعة الأدمن.',
    )
    conn.close()
    return

  try:
    # إنشاء رابط دعوة مخصص استخدام واحد فقط (Single-use) وينتهي بعد 10 دقائق
    invite_link = bot.create_chat_invite_link(
        chat_id=chat_target_id,
        member_limit=1,
        expire_date=int(message.date) + 600,
    )

    # 🔒 قفل الحساب وتحديث الحالة إلى joined = 1 فوراً
    cursor.execute(
        """
            UPDATE students 
            SET telegram_user_id = ?, joined = 1 
            WHERE student_id = ?
        """,
        (user_id, student_id),
    )
    conn.commit()

    success_msg = (
        f'✅ **تم التحقق من بياناتك بنجاح!**\n\n'
        f'👤 **الاسم:** {name}\n'
        f'🆔 **الرقم الجامعي:** {student_id}\n'
        f'📚 **السنة الدراسية:** {year_name}\n\n'
        f'🔗 **رابط الانضمام الخاطف (صالح لاستخدام شخص واحد ولمرة واحدة'
        f' فقط):**\n'
        f'{invite_link.invite_link}'
    )
    bot.send_message(
        message.chat.id,
        success_msg,
        reply_markup=types.ReplyKeyboardRemove(),
        parse_mode='Markdown',
    )

  except Exception as e:
    bot.send_message(
        message.chat.id, f'❌ حدث خطأ أثناء إنشاء رابط الانضمام: {str(e)}'
    )

  conn.close()


# ==================== لوحة إحصائيات الأدمن (/stats) ====================
@bot.message_handler(commands=['stats'])
def admin_stats(message):
  if message.from_user.id != ADMIN_ID:
    return

  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()

  cursor.execute('SELECT COUNT(*) FROM students')
  total_students = cursor.fetchone()[0]

  cursor.execute('SELECT COUNT(*) FROM students WHERE joined = 1')
  joined_students = cursor.fetchone()[0]

  cursor.execute(
      'SELECT year, COUNT(*), SUM(joined) FROM students GROUP BY year'
  )
  year_stats = cursor.fetchall()
  conn.close()

  text = '📊 **إحصائيات الانتقال للتلغرام:**\n\n'
  text += f'👥 إجمالي الطلاب المسجلين: **{total_students}**\n'
  text += f'✅ إجمالي المنضمين فعلياً: **{joined_students}**\n\n'
  text += '📌 **التفاصيل حسب السنة:**\n'

  for y_name, count, joined_count in year_stats:
    joined_c = joined_count if joined_count else 0
    text += f'• {y_name}: **{joined_c} / {count}** انضموا\n'

  bot.send_message(message.chat.id, text, parse_mode='Markdown')


# تشغيل البوت بشكل مستمر
bot.infinity_polling()

