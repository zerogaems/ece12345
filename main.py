import io
import os
import re
import sqlite3
from threading import Thread
from flask import Flask
import pandas as pd
import telebot
from telebot import types

# ==================== خادم Flask لإرضاء Render و UptimeRobot ====================
app = Flask('')


@app.route('/')
def home():
  return 'Bot is running smoothly with Dual Links on Render!'


def run_flask():
  port = int(os.environ.get('PORT', 8080))
  app.run(host='0.0.0.0', port=port)


# ==================== الإعدادات الأساسية ====================
BOT_TOKEN = os.environ.get(
    'BOT_TOKEN', 'ضع_التوكن_هنا_إن_لم_تستخدم_متغيرات_البيئة'
)
ADMIN_ID = int(os.environ.get('ADMIN_ID', '7547218555'))

bot = telebot.TeleBot(BOT_TOKEN)

# معرفات القنوات والمجموعات لكل سنة دراسية (Chat IDs)
# ملاحظة: أضف البوت كمشرف (Admin) في جميع القنوات والمجموعات مع إعطائه صلاحية "Invite via Link"
YEAR_CHATS = {
    'السنة الأولى': {
        'lectures': -1004413316628,  # ID قناة المحاضرات
        'discussion': -1003953300954,  # ID مجموعة المناقشة
    },
    'السنة الثانية': {
        'lectures': -1004372697822,
        'discussion': -1004391353751,
    },
    'السنة الثالثة': {
        'lectures': -1004449351242,
        'discussion': -1004495552551,
    },
    'السنة الرابعة': {
        'lectures': -1004382433468,
        'discussion': -1004364259833,
    },
    'السنة الخامسة': {
        'lectures': -1003852549374,
        'discussion': -1004363731175,
    },
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


# ==================== تنظيف وتوحيد أرقام الهواتف ====================
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


# ==================== معالجة واستيراد ملف Excel ====================
def process_excel_file(file_path):
  df = pd.read_excel(file_path)
  df = df.iloc[:, [1, 2, 3, 4]]
  df.columns = ['Name', 'Student_ID', 'Year', 'Phone']

  df['Student_ID'] = df['Student_ID'].astype(str).str.strip()
  df['Phone'] = df['Phone'].apply(clean_phone)
  df['Year'] = df['Year'].astype(str).str.strip()

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
  return added_count


if os.path.exists('uploaded_responses.xlsx'):
  try:
    process_excel_file('uploaded_responses.xlsx')
  except Exception:
    pass


# ==================== رفع واستيراد Excel (للأدمن) ====================
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

    added_count = process_excel_file(excel_path)

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


# ==================== الأوامر الإدارية للأدمن ====================


@bot.message_handler(commands=['search'])
def search_student(message):
  if message.from_user.id != ADMIN_ID:
    return
  args = message.text.split()
  if len(args) < 2:
    bot.reply_to(
        message,
        '⚠️ يرجى إدخال الرقم الجامعي:\n`/search 123456`',
        parse_mode='Markdown',
    )
    return

  sid = args[1].strip()
  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()
  cursor.execute(
      'SELECT name, year, phone, joined, telegram_user_id, joined_at FROM'
      ' students WHERE student_id = ?',
      (sid,),
  )
  st = cursor.fetchone()
  conn.close()

  if not st:
    bot.reply_to(message, '❌ لم يتم العثور على هذا الرقم الجامعي.')
    return

  name, year, phone, joined, tg_id, joined_at = st
  status = '✅ انضم بالفعل' if joined == 1 else '❌ لم ينضم بعد'

  res = (
      f'🔍 **بيانات الطالب:**\n\n'
      f'👤 **الاسم:** {name}\n'
      f'🆔 **الرقم الجامعي:** {sid}\n'
      f'📚 **السنة:** {year}\n'
      f'📱 **الرقم:** `{phone}`\n'
      f'📌 **الحالة:** {status}\n'
      f'🆔 **Telegram ID:** `{tg_id if tg_id else "غير مسجل"}`\n'
      f'🕒 **تاريخ الانضمام:** {joined_at}'
  )
  bot.reply_to(message, res, parse_mode='Markdown')


@bot.message_handler(commands=['update_phone'])
def update_phone_cmd(message):
  if message.from_user.id != ADMIN_ID:
    return
  args = message.text.split()
  if len(args) < 3:
    bot.reply_to(
        message,
        '⚠️ الاستخدام الصحيح:\n`/update_phone الرقم_الجامعي الرقم_الجديد`',
        parse_mode='Markdown',
    )
    return

  sid = args[1].strip()
  new_phone = clean_phone(args[2].strip())

  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()
  cursor.execute(
      'UPDATE students SET phone = ? WHERE student_id = ?', (new_phone, sid)
  )

  if cursor.rowcount > 0:
    conn.commit()
    bot.reply_to(
        message,
        f'✅ تم تحديث رقم الطالب ({sid}) إلى `{new_phone}` بنجاح.',
        parse_mode='Markdown',
    )
  else:
    bot.reply_to(message, '❌ لم يتم العثور على هذا الرقم الجامعي.')
  conn.close()


@bot.message_handler(commands=['export'])
def export_data(message):
  if message.from_user.id != ADMIN_ID:
    return
  bot.reply_to(message, '⏳ جاري استخراج ملف البيانات الكامل...')

  conn = sqlite3.connect('telecom_students.db')
  df = pd.read_sql_query('SELECT * FROM students', conn)
  conn.close()

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Students_Status')
  output.seek(0)

  bot.send_document(
      message.chat.id,
      document=types.InputFile(output, filename='Telecom_Students_Report.xlsx'),
      caption='📊 **تقرير الطلاب الكامل وحالات الانضمام**',
      parse_mode='Markdown',
  )


@bot.message_handler(commands=['broadcast'])
def broadcast_msg(message):
  if message.from_user.id != ADMIN_ID:
    return
  msg_text = message.text.replace('/broadcast', '').strip()
  if not msg_text:
    bot.reply_to(
        message,
        '⚠️ اكتب الرسالة بعد الأمر. مثال:\n`/broadcast السلام عليكم`',
        parse_mode='Markdown',
    )
    return

  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()
  cursor.execute(
      'SELECT telegram_user_id FROM students WHERE joined = 1 AND'
      ' telegram_user_id IS NOT NULL'
  )
  users = cursor.fetchall()
  conn.close()

  success = 0
  for u in users:
    try:
      bot.send_message(
          u[0], f'📢 **إعلان من هيئة الاتصالات:**\n\n{msg_text}', parse_mode='Markdown'
      )
      success += 1
    except Exception:
      continue

  bot.reply_to(
      message,
      f'✅ تم إرسال الإعلان بنجاح إلى **{success}** طالب مسجل.',
      parse_mode='Markdown',
  )


@bot.message_handler(commands=['reset'])
def reset_student(message):
  if message.from_user.id != ADMIN_ID:
    return
  args = message.text.split()
  if len(args) < 2:
    bot.reply_to(
        message,
        '⚠️ اكتب الرقم الجامعي:\n`/reset 123456`',
        parse_mode='Markdown',
    )
    return

  sid = args[1].strip()
  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()
  cursor.execute(
      'UPDATE students SET telegram_user_id = NULL, joined = 0 WHERE'
      ' student_id = ?',
      (sid,),
  )

  if cursor.rowcount > 0:
    conn.commit()
    bot.reply_to(message, f'✅ تم فك القفل عن الرقم الجامعي ({sid}).')
  else:
    bot.reply_to(message, '❌ الرقم الجامعي غير موجود.')
  conn.close()


# ==================== تفاعل الطلاب (/start) ====================
@bot.message_handler(commands=['start'])
def start_command(message):
  user_id = message.from_user.id

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
        f'لقد حصلت على روابط الانضمام لسنتك الدراسية (**{existing_user_by_id[1]}**) سابقاً.\n'
        f'⚠️ **النظام يمنع الحصول على روابط أخرى.**',
        parse_mode='Markdown',
    )
    conn.close()
    return

  conn.close()

  markup = types.ReplyKeyboardMarkup(
      row_width=1, resize_keyboard=True, one_time_keyboard=True
  )
  button = types.KeyboardButton(
      text='📱 مشاركة رقم التليغرام لتأكيد الهوية', request_contact=True
  )
  markup.add(button)

  welcome_text = (
      'أهلاً بك في البوت الرسمي لقسم الهندسة الإلكترونية والاتصالات (الهمك)'
      ' 🎓\n\nللحصول على روابط مجموعة المحاضرات ومجموعة المناقشة الخاصة بسنتك'
      ' الدراسية، يرجى الضغط على الزر أدناه لمشاركة رقمك المعتمد في الاستبيان.'
  )
  bot.send_message(message.chat.id, welcome_text, reply_markup=markup)


# ==================== استقبال جهة الاتصال والتوزيع بالرابطين ====================
@bot.message_handler(content_types=['contact'])
def handle_contact(message):
  if not message.contact:
    return

  raw_phone = message.contact.phone_number
  cleaned_phone = clean_phone(raw_phone)
  user_id = message.from_user.id

  conn = sqlite3.connect('telecom_students.db')
  cursor = conn.cursor()

  cursor.execute(
      'SELECT name, year, joined FROM students WHERE telegram_user_id = ?',
      (user_id,),
  )
  already_joined_account = cursor.fetchone()
  if already_joined_account and already_joined_account[2] == 1:
    bot.send_message(
        message.chat.id,
        '🚫 **عذراً! حسابك التليغرام مسجّل ومستلم للروابط سابقاً.**',
        reply_markup=types.ReplyKeyboardRemove(),
        parse_mode='Markdown',
    )
    conn.close()
    return

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

  if joined == 1:
    bot.send_message(
        message.chat.id,
        f'⚠️ **عذراً، هذا الرقم الجامعي ({student_id}) استلم روابط الانضمام'
        ' سابقاً ولا يمكن استخدامه مجدداً.**',
        reply_markup=types.ReplyKeyboardRemove(),
        parse_mode='Markdown',
    )
    conn.close()
    return

  year_data = YEAR_CHATS.get(year_name)

  if not year_data:
    bot.send_message(
        message.chat.id,
        '⚠️ خطأ في إعدادات السنة الدراسية، يرجى مراجعة الأدمن.',
    )
    conn.close()
    return

  try:
    # 1. إنشاء رابط دعوة شخصي لقناة المحاضرات (استخدام واحد)
    lectures_link = bot.create_chat_invite_link(
        chat_id=year_data['lectures'],
        member_limit=1,
        expire_date=int(message.date) + 600,
    ).invite_link

    # 2. إنشاء رابط دعوة شخصي لمجموعة المناقشة (استخدام واحد)
    discussion_link = bot.create_chat_invite_link(
        chat_id=year_data['discussion'],
        member_limit=1,
        expire_date=int(message.date) + 600,
    ).invite_link

    # قفل الحساب وتحديث الحالة كـ joined = 1
    cursor.execute(
        """
            UPDATE students 
            SET telegram_user_id = ?, joined = 1 
            WHERE student_id = ?
        """,
        (user_id, student_id),
    )
    conn.commit()

    # أزرار انضمام مباشرة تحت الرسالة
    markup = types.InlineKeyboardMarkup(row_width=1)
    btn1 = types.InlineKeyboardButton(
        text='📚 الانضمام لقناة المحاضرات', url=lectures_link
    )
    btn2 = types.InlineKeyboardButton(
        text='💬 الانضمام لمجموعة المناقشة', url=discussion_link
    )
    markup.add(btn1, btn2)

    success_msg = (
        f'✅ **تم التحقق من بياناتك بنجاح!**\n\n'
        f'👤 **الاسم:** {name}\n'
        f'🆔 **الرقم الجامعي:** {student_id}\n'
        f'📚 **السنة الدراسية:** {year_name}\n\n'
        f'👇 **إليك روابط الانضمام الرسمية المخصصة لسنتك:**\n'
        f'*(ملاحظة: الروابط شخصية وخاصة بك وصالحة للاستخدام مرة واحدة فقط)*'
    )

    bot.send_message(
        message.chat.id,
        success_msg,
        reply_markup=markup,
        parse_mode='Markdown',
    )

  except Exception as e:
    bot.send_message(
        message.chat.id, f'❌ حدث خطأ أثناء إنشاء روابط الانضمام: {str(e)}'
    )

  conn.close()


# ==================== الإحصائيات (/stats) ====================
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


# ==================== التشغيل ====================
def run_bot():
  bot.infinity_polling(skip_pending=True)


if __name__ == '__main__':
  bot_thread = Thread(target=run_bot)
  bot_thread.daemon = True
  bot_thread.start()

  run_flask()

