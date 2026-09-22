import io
import logging
import os
import re
import time
from threading import Thread
from flask import Flask
import libsql
import pandas as pd
import telebot
from telebot import types

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
)
logger = logging.getLogger('telecom_bot')

# ==================== خادم Flask لإرضاء Render و UptimeRobot ====================
app = Flask('')


@app.route('/')
def home():
  return 'Bot is running smoothly with Admin Panel on Render!'


def run_flask():
  port = int(os.environ.get('PORT', 8080))
  app.run(host='0.0.0.0', port=port)


# ==================== الإعدادات الأساسية ====================
BOT_TOKEN = os.environ.get('BOT_TOKEN')
if not BOT_TOKEN:
  raise RuntimeError('يجب ضبط متغير البيئة BOT_TOKEN قبل تشغيل البوت.')

ADMIN_ID = int(os.environ.get('ADMIN_ID', '123456789'))

# معرف الأدمن/ممثل الهيئة على تيليجرام (بدون @) - يظهر للطلاب عند الحاجة للتواصل
ADMIN_USERNAME = 'Youssef_Sabra'

bot = telebot.TeleBot(BOT_TOKEN)

# ==================== الاتصال بقاعدة بيانات Turso (بدل ملف SQLite محلي) ====================
# بدل ما نخزن الملف محلياً (بينمسح مع كل إعادة تشغيل على Render)، نتصل بقاعدة
# بيانات Turso السحابية الدائمة. لازم تضبط هدين المتغيرين على Render:
# TURSO_DATABASE_URL و TURSO_AUTH_TOKEN
TURSO_URL = os.environ.get('TURSO_DATABASE_URL')
TURSO_AUTH_TOKEN = os.environ.get('TURSO_AUTH_TOKEN')

if not TURSO_URL or not TURSO_AUTH_TOKEN:
  raise RuntimeError(
      'يجب ضبط TURSO_DATABASE_URL و TURSO_AUTH_TOKEN كمتغيرات بيئة قبل'
      ' التشغيل.'
  )


def get_connection():
  """يفتح اتصال جديد بقاعدة بيانات Turso. الواجهة نفس sqlite3 تقريباً
  (cursor / execute / commit / close) لذلك باقي الكود ما احتاج تعديل كبير."""
  return libsql.connect(database=TURSO_URL, auth_token=TURSO_AUTH_TOKEN)


# قاموس حفظ حالات الإدخال المؤقتة للأدمن
user_states = {}

# معرفات القنوات والمجموعات لكل سنة دراسية (Chat IDs)
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
  conn = get_connection()
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


# ==================== تنظيف وتوحيد أرقام الهواتف والسنوات ====================
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


def normalize_year(year_str):
  y = str(year_str).strip()
  if 'أول' in y or 'اول' in y or '1' in y:
    return 'السنة الأولى'
  if 'ثان' in y or '2' in y:
    return 'السنة الثانية'
  if 'ثالث' in y or '3' in y:
    return 'السنة الثالثة'
  if 'رابع' in y or '4' in y:
    return 'السنة الرابعة'
  if 'خامس' in y or '5' in y:
    return 'السنة الخامسة'
  return y


# ==================== معالجة واستيراد ملف Excel ====================
def find_column(df, keywords, label):
  col = next((c for c in df.columns if any(k in c for k in keywords)), None)
  if col is None:
    raise ValueError(f'لم يتم العثور على عمود "{label}" في ملف الإكسل.')
  return col


def process_excel_file(file_path):
  df = pd.read_excel(file_path)
  df.columns = [str(c).strip() for c in df.columns]

  # بحث صارم بالاسم عن كل عمود مطلوب - إذا لم يوجد نرفض الاستيراد
  # بدل الرجوع لعمود عشوائي بالموضع (كان يسبب استيراد بيانات خاطئة بصمت)
  name_col = find_column(df, ['اسم', 'الاسم', 'Name'], 'الاسم')
  id_col = find_column(
      df, ['جامعي', 'رقم', 'ID', 'Student'], 'الرقم الجامعي'
  )
  year_col = find_column(df, ['سنة', 'السنة', 'Year'], 'السنة الدراسية')
  phone_col = find_column(
      df, ['هاتف', 'موبايل', 'واتس', 'Phone'], 'رقم الهاتف'
  )

  df_students = pd.DataFrame({
      'Name': df[name_col].astype(str).str.strip(),
      'Student_ID': df[id_col].astype(str).str.strip(),
      'Year': df[year_col].apply(normalize_year),
      'Phone': df[phone_col].apply(clean_phone),
  })

  df_clean = df_students.drop_duplicates(subset=['Student_ID'], keep='last')

  conn = get_connection()
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
    except Exception as e:
      logger.warning(
          'تعذر استيراد الطالب %s: %s', row['Student_ID'], e
      )
      continue

  conn.commit()
  conn.close()
  return added_count


if os.path.exists('uploaded_responses.xlsx'):
  try:
    process_excel_file('uploaded_responses.xlsx')
  except Exception as e:
    logger.warning('فشل استيراد ملف الإكسل المحفوظ سابقاً: %s', e)


# ==================== توليد ملف الطلاب (الكل / المتبقين فقط) ====================
def build_students_excel(remaining_only=False):
  """يبني ملف Excel لبيانات الطلاب.
  remaining_only=True: يستثني الطلاب الذين انضموا فعلاً (joined=1)،
  بحيث يمثل الملف قائمة "الباقي" فقط - وهذا يعني أن أي طالب ينضم
  يختفي تلقائياً من هذا الملف بمجرد إعادة توليده.
  """
  conn = get_connection()
  cursor = conn.cursor()
  if remaining_only:
    cursor.execute('SELECT * FROM students WHERE joined = 0')
  else:
    cursor.execute('SELECT * FROM students')
  rows = cursor.fetchall()
  columns = [d[0] for d in cursor.description]
  df = pd.DataFrame(rows, columns=columns)
  conn.close()

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine='openpyxl') as writer:
    sheet = 'Remaining_Students' if remaining_only else 'Students_Status'
    df.to_excel(writer, index=False, sheet_name=sheet)
  output.seek(0)
  return output, len(df)


# ==================== لوحة التحكم للأدمن (Admin Panel) ====================
def get_admin_keyboard():
  markup = types.InlineKeyboardMarkup(row_width=2)
  btn_stats = types.InlineKeyboardButton(
      text='📊 الإحصائيات العامة', callback_data='admin_stats'
  )
  btn_export = types.InlineKeyboardButton(
      text='📥 تنزيل تقرير Excel (الكل)', callback_data='admin_export'
  )
  btn_export_remaining = types.InlineKeyboardButton(
      text='📄 تنزيل ملف المتبقين (لم ينضموا)',
      callback_data='admin_export_remaining',
  )
  btn_search = types.InlineKeyboardButton(
      text='🔍 البحث عن طالب', callback_data='admin_search'
  )
  btn_reset = types.InlineKeyboardButton(
      text='🔓 فك قفل طالب', callback_data='admin_reset'
  )
  btn_broadcast = types.InlineKeyboardButton(
      text='📢 إرسال إعلان للجميع', callback_data='admin_broadcast'
  )
  btn_update_phone = types.InlineKeyboardButton(
      text='📱 تعديل رقم طالب', callback_data='admin_update_phone'
  )

  markup.add(btn_stats, btn_export)
  markup.add(btn_export_remaining)
  markup.add(btn_search, btn_reset)
  markup.add(btn_update_phone, btn_broadcast)
  return markup


@bot.message_handler(commands=['admin'])
def admin_command(message):
  if message.from_user.id != ADMIN_ID:
    bot.reply_to(message, '⚠️ عذراً، هذه اللوحة مخصصة لرئيس الهيئة/الأدمن فقط.')
    return

  text = (
      '🛠️ **لوحة تحكم رئيس الهيئة (Admin Panel)**\n\n'
      'مرحباً بك! اختر الخدمة التي تريد تنفيذها من الأزرار أدناه، أو أرسل ملف'
      ' الـ Excel لتحديث البيانات فوراً.'
  )
  bot.send_message(
      message.chat.id,
      text,
      reply_markup=get_admin_keyboard(),
      parse_mode='Markdown',
  )


# ==================== التفاعل مع أزرار لوحة الأدمن (Callback Query) ====================
@bot.callback_query_handler(func=lambda call: call.data.startswith('admin_'))
def handle_admin_callbacks(call):
  if call.from_user.id != ADMIN_ID:
    bot.answer_callback_query(call.id, '⚠️ غير مصرح لك.', show_alert=True)
    return

  action = call.data

  if action == 'admin_stats':
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) FROM students')
    total = cursor.fetchone()[0]
    cursor.execute('SELECT COUNT(*) FROM students WHERE joined = 1')
    joined = cursor.fetchone()[0]
    cursor.execute(
        'SELECT year, COUNT(*), SUM(joined) FROM students GROUP BY year'
    )
    year_stats = cursor.fetchall()
    conn.close()

    text = '📊 **إحصائيات الانتقال للتلغرام:**\n\n'
    text += f'👥 إجمالي المسجلين بالاستبيان: **{total}**\n'
    text += f'✅ إجمالي المنضمين فعلياً: **{joined}**\n\n'
    text += '📌 **التفاصيل حسب السنة:**\n'

    for y_name, count, joined_count in year_stats:
      jc = joined_count if joined_count else 0
      text += f'• {y_name}: **{jc} / {count}** انضموا\n'

    bot.send_message(
        call.message.chat.id,
        text,
        reply_markup=get_admin_keyboard(),
        parse_mode='Markdown',
    )
    bot.answer_callback_query(call.id)

  elif action == 'admin_export':
    bot.answer_callback_query(call.id, '⏳ جاري جلب ملف البيانات...')
    output, count = build_students_excel(remaining_only=False)

    bot.send_document(
        call.message.chat.id,
        document=types.InputFile(
            output, filename='Telecom_Students_Report.xlsx'
        ),
        caption=f'📊 **تقرير الطلاب الكامل وحالات الانضمام ({count} طالب)**',
        parse_mode='Markdown',
    )

  elif action == 'admin_export_remaining':
    bot.answer_callback_query(call.id, '⏳ جاري تجهيز ملف المتبقين...')
    output, count = build_students_excel(remaining_only=True)

    bot.send_document(
        call.message.chat.id,
        document=types.InputFile(
            output, filename='Remaining_Students.xlsx'
        ),
        caption=(
            f'📄 **ملف الطلاب الذين لم ينضموا بعد ({count} طالب)**\n'
            'ملاحظة: أي طالب ينضم يُحذف تلقائياً من هذا الملف عند إعادة تنزيله.'
        ),
        parse_mode='Markdown',
    )

  elif action == 'admin_search':
    user_states[call.from_user.id] = 'awaiting_search_id'
    bot.send_message(
        call.message.chat.id,
        '🔍 **يرجى إرسال الرقم الجامعي للطالب المراد البحث عنه:**',
    )
    bot.answer_callback_query(call.id)

  elif action == 'admin_reset':
    user_states[call.from_user.id] = 'awaiting_reset_id'
    bot.send_message(
        call.message.chat.id,
        '🔓 **يرجى إرسال الرقم الجامعي للطالب المراد فك قفله:**',
    )
    bot.answer_callback_query(call.id)

  elif action == 'admin_broadcast':
    user_states[call.from_user.id] = 'awaiting_broadcast_msg'
    bot.send_message(
        call.message.chat.id,
        '📢 **اكتب الرسالة التي تريد بثها لجميع الطلاب المنضمين:**',
    )
    bot.answer_callback_query(call.id)

  elif action == 'admin_update_phone':
    user_states[call.from_user.id] = 'awaiting_update_phone'
    bot.send_message(
        call.message.chat.id,
        '📱 **أرسل الرقم الجامعي والرقم الجديد مفصولين بمسافة**\nمثال:'
        ' `123456 0912345678`:',
        parse_mode='Markdown',
    )
    bot.answer_callback_query(call.id)


# ==================== استقبال إدخالات الأدمن التفاعلية ====================
@bot.message_handler(
    func=lambda msg: msg.from_user.id == ADMIN_ID
    and msg.from_user.id in user_states
)
def handle_admin_inputs(message):
  state = user_states.get(message.from_user.id)

  if state == 'awaiting_search_id':
    sid = message.text.strip()
    conn = get_connection()
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
    else:
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

    del user_states[message.from_user.id]

  elif state == 'awaiting_reset_id':
    sid = message.text.strip()
    conn = get_connection()
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
    del user_states[message.from_user.id]

  elif state == 'awaiting_broadcast_msg':
    msg_text = message.text.strip()
    conn = get_connection()
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
            u[0],
            f'📢 **إعلان من هيئة الاتصالات:**\n\n{msg_text}',
            parse_mode='Markdown',
        )
        success += 1
      except Exception as e:
        logger.warning('فشل إرسال الإعلان للمستخدم %s: %s', u[0], e)
        continue
      # تهدئة بسيطة لتجنب حد الإرسال (Flood Control) من تيليجرام
      time.sleep(0.05)

    bot.reply_to(
        message,
        f'✅ تم إرسال الإعلان بنجاح إلى **{success}** طالب.',
        parse_mode='Markdown',
    )
    del user_states[message.from_user.id]

  elif state == 'awaiting_update_phone':
    args = message.text.split()
    if len(args) < 2:
      bot.reply_to(
          message,
          '⚠️ صيغة غير صحيحة. أرسل الرقم الجامعي ثم الهاتف المحدث مع مسافة'
          ' بينهما.',
      )
    else:
      sid = args[0].strip()
      new_phone = clean_phone(args[1].strip())
      conn = get_connection()
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
        bot.reply_to(message, '❌ الرقم الجامعي غير موجود.')
      conn.close()
    del user_states[message.from_user.id]


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
    logger.exception('فشل معالجة ملف الإكسل المرفوع')
    bot.edit_message_text(
        f'❌ حدث خطأ أثناء معالجة الملف: {str(e)}',
        chat_id=message.chat.id,
        message_id=msg.message_id,
    )


# ==================== تفاعل الطلاب (/start) ====================
@bot.message_handler(commands=['start'])
def start_command(message):
  user_id = message.from_user.id

  # إذا كان المستخدم هو الأدمن، تعرض له اللوحة مباشرة
  if user_id == ADMIN_ID:
    admin_command(message)
    return

  conn = get_connection()
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
        f'لقدحصلت على روابط الانضمام لسنتك الدراسية (**{existing_user_by_id[1]}**) سابقاً.\n'
        f'⚠️ **النظام يمنع الحصول على روابط أخرى.**\n\n'
        f'📞 لأي استفسار تواصل مع الهيئة: @{ADMIN_USERNAME}',
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
      f'\n\n📞 لأي استفسار أو مشكلة تواصل مع الهيئة: @{ADMIN_USERNAME}'
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

  conn = get_connection()
  cursor = conn.cursor()

  cursor.execute(
      'SELECT name, year, joined FROM students WHERE telegram_user_id = ?',
      (user_id,),
  )
  already_joined_account = cursor.fetchone()
  if already_joined_account and already_joined_account[2] == 1:
    bot.send_message(
        message.chat.id,
        '🚫 **عذراً! حسابك التليغرام مسجّل ومستلم للروابط سابقاً.**\n\n'
        f'📞 لأي استفسار تواصل مع الهيئة: @{ADMIN_USERNAME}',
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
        f' لمراجعة بياناتك: @{ADMIN_USERNAME}',
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
        ' سابقاً ولا يمكن استخدامه مجدداً.**\n\n'
        f'📞 لأي استفسار تواصل مع الهيئة: @{ADMIN_USERNAME}',
        reply_markup=types.ReplyKeyboardRemove(),
        parse_mode='Markdown',
    )
    conn.close()
    return

  year_data = YEAR_CHATS.get(year_name)

  if not year_data:
    bot.send_message(
        message.chat.id,
        '⚠️ خطأ في إعدادات السنة الدراسية، يرجى مراجعة الأدمن:'
        f' @{ADMIN_USERNAME}',
    )
    conn.close()
    return

  try:
    lectures_link = bot.create_chat_invite_link(
        chat_id=year_data['lectures'],
        member_limit=1,
        expire_date=int(message.date) + 600,
    ).invite_link

    discussion_link = bot.create_chat_invite_link(
        chat_id=year_data['discussion'],
        member_limit=1,
        expire_date=int(message.date) + 600,
    ).invite_link

    cursor.execute(
        """
            UPDATE students 
            SET telegram_user_id = ?, joined = 1 
            WHERE student_id = ?
        """,
        (user_id, student_id),
    )
    conn.commit()

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

    # إشعار الأدمن (اختياري) أن طالباً جديداً انضم - مفيد لتتبع التقدم لحظياً
    try:
      bot.send_message(
          ADMIN_ID,
          f'ℹ️ الطالب **{name}** ({student_id} - {year_name}) انضم الآن.\n'
          'تم حذفه من قائمة "المتبقين" تلقائياً.',
          parse_mode='Markdown',
      )
    except Exception as e:
      logger.warning('تعذر إشعار الأدمن بانضمام طالب: %s', e)

  except Exception as e:
    logger.exception('فشل إنشاء روابط الانضمام')
    bot.send_message(
        message.chat.id, f'❌ حدث خطأ أثناء إنشاء روابط الانضمام: {str(e)}'
    )

  conn.close()


# ==================== التشغيل ====================
def run_bot():
  bot.infinity_polling(skip_pending=True)


if __name__ == '__main__':
  bot_thread = Thread(target=run_bot)
  bot_thread.daemon = True
  bot_thread.start()

  run_flask()
