import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
import openpyxl
import os
import re
import html

# Apnar 2nd Bot Token
BOT_TOKEN = "8997251206:AAEkNWjTPMi_3xN-pDDjMIGNO2SWe9qy01E"
bot = telebot.TeleBot(BOT_TOKEN)

user_sessions = {}

def get_navigation_buttons(current_index, total_count):
    markup = InlineKeyboardMarkup(row_width=2)
    nav_buttons = []
    
    if current_index > 0:
        nav_buttons.append(InlineKeyboardButton("⬅️ Previous", callback_data="prev_item"))
        
    if current_index < total_count - 1:
        nav_buttons.append(InlineKeyboardButton("Next ➡️", callback_data="next_item"))
        
    if nav_buttons:
        markup.add(*nav_buttons)
        
    markup.add(InlineKeyboardButton("🔄 Upload New File", callback_data="reset_bot"))
    return markup

def format_account_message(index, total, item):
    uid = html.escape(str(item['uid']))
    password = html.escape(str(item['password']))
    cookie = html.escape(str(item['cookie']))
    
    text = (
        f"📊 <b>Account No: {index + 1} / {total}</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n\n"
        f"🆔 <b>UID (Tap to copy):</b>\n"
        f"<code>{uid}</code>\n\n"
        f"🔑 <b>Password (Tap to copy):</b>\n"
        f"<code>{password}</code>\n\n"
        f"🍪 <b>Cookie (Tap to copy):</b>\n"
        f"<code>{cookie}</code>\n\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"💡 <i>Box-er lekhar upor 1 bar tap korlei copy hoye jabe.</i>"
    )
    return text

@bot.message_handler(commands=['start'])
def start_cmd(message):
    chat_id = message.chat.id
    user_sessions.pop(chat_id, None)
    
    welcome_text = (
        "👋 <b>Welcome! (Excel Reader & Dead ID Filter Bot)</b>\n\n"
        "1️⃣ Prothome apnar <b>.xlsx (Excel)</b> file-ti upload korun.\n"
        "2️⃣ Erpor bot apnar kache <b>Dead UID</b> chaibe (na thakle <code>nothing</code> likhben).\n"
        "3️⃣ Bot Dead ID gulo bad diye baki sob Live ID ebong Cookie একটার পর একটা Tap-to-Copy আকারে দেখাবে!"
    )
    bot.send_message(chat_id, welcome_text, parse_mode="HTML")

@bot.message_handler(content_types=['document'])
def handle_excel_file(message):
    chat_id = message.chat.id
    file_name = message.document.file_name

    if not file_name.lower().endswith('.xlsx'):
        bot.reply_to(message, "❌ Sudhumatro <b>.xlsx (Excel)</b> file upload korun!", parse_mode="HTML")
        return

    wait_msg = bot.reply_to(message, "⏳ <b>File load kora hocche, ektu opekha korun...</b>", parse_mode="HTML")

    file_path = f"session_{chat_id}.xlsx"
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        with open(file_path, 'wb') as f:
            f.write(downloaded_file)

        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active

        raw_accounts = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 3:
                continue
            uid, password, cookie = row[0], row[1], row[2]
            if uid is None and cookie is None:
                continue
            raw_accounts.append({
                'uid': str(uid).strip() if uid else "N/A",
                'password': str(password).strip() if password else "N/A",
                'cookie': str(cookie).strip() if cookie else "N/A"
            })

        wb.close()
        if os.path.exists(file_path):
            os.remove(file_path)

        if not raw_accounts:
            bot.edit_message_text("⚠️ Ei Excel file-tite kono data pawa jayni!", chat_id, wait_msg.message_id)
            return

        # Data save kore Dead UID chawa hocche
        user_sessions[chat_id] = {
            'step': 'awaiting_dead_uids',
            'raw_items': raw_accounts,
            'items': [],
            'index': 0
        }

        ask_dead_text = (
            f"✅ <b>File সফলভাবে লোড হয়েছে!</b>\n"
            f"📑 মোট অ্যাকাউন্ট পাওয়া গেছে: <b>{len(raw_accounts)} টি</b>\n\n"
            f"🗑️ <b>এখন আপনার কাছে কোনো Dead UID থাকলে সেগুলো মেসেজ করে পাঠান:</b>\n"
            f"<i>(একাধিক UID থাকলে নিচে নিচে বা স্পেস/কমা দিয়ে একসাথে পাঠাতে পারেন। ওই UID গুলো বাদ দেওয়া হবে।)</i>\n\n"
            f"👉 আর যদি কোনো Dead ID না থাকে, তবে <code>nothing</code> লিখে পাঠান।"
        )

        bot.edit_message_text(
            ask_dead_text,
            chat_id,
            wait_msg.message_id,
            parse_mode="HTML"
        )

    except Exception as e:
        if os.path.exists(file_path):
            os.remove(file_path)
        bot.edit_message_text(f"❌ File read korte somossa hoyeche: {e}", chat_id, wait_msg.message_id)

@bot.message_handler(content_types=['text'], func=lambda msg: user_sessions.get(msg.chat.id, {}).get('step') == 'awaiting_dead_uids')
def filter_dead_uids_and_start(message):
    chat_id = message.chat.id
    user_input = message.text.strip()
    session = user_sessions.get(chat_id)

    if not session or 'raw_items' not in session:
        bot.reply_to(message, "⚠️ Session expire hoye geche! Notun kore .xlsx file upload korun.")
        return

    raw_items = session['raw_items']

    # Jodi user "nothing" lekhe
    if user_input.lower() == 'nothing':
        dead_uids_set = set()
    else:
        # Message theke sob number/UID alada kora (comma, space ba newline jai thakuk)
        extracted_ids = re.findall(r'\d+', user_input)
        dead_uids_set = set(extracted_ids)

    # Dead UID gulo bad diye final list toiri kora
    filtered_items = [
        item for item in raw_items 
        if str(item['uid']).strip() not in dead_uids_set
    ]

    removed_count = len(raw_items) - len(filtered_items)

    if not filtered_items:
        bot.reply_to(
            message,
            f"⚠️ সব অ্যাকাউন্টই বাদ পড়ে গেছে! (মোট বাদ পড়েছে: {removed_count} টি)।\nনতুন ফাইল আপলোড করতে /start দিন।"
        )
        user_sessions.pop(chat_id, None)
        return

    # Session update kora
    user_sessions[chat_id]['step'] = 'viewing'
    user_sessions[chat_id]['items'] = filtered_items
    user_sessions[chat_id]['index'] = 0
    user_sessions[chat_id].pop('raw_items', None)

    total = len(filtered_items)
    first_item = filtered_items[0]

    summary_header = (
        f"✅ <b>Filter Complete!</b>\n"
        f"❌ বাদ দেওয়া হয়েছে (Dead): <b>{removed_count} টি</b>\n"
        f"✔️ বাকি আছে (Active): <b>{total} টি</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
    )

    msg_text = summary_header + format_account_message(0, total, first_item)
    if total == 1:
        msg_text += "\n\n🎉 <b>এটিই তালিকার একমাত্র অ্যাকাউন্ট!</b>"

    bot.send_message(
        chat_id,
        msg_text,
        parse_mode="HTML",
        reply_markup=get_navigation_buttons(0, total)
    )

@bot.callback_query_handler(func=lambda call: call.data in ["next_item", "prev_item", "reset_bot"])
def handle_navigation(call):
    chat_id = call.message.chat.id

    if call.data == "reset_bot":
        user_sessions.pop(chat_id, None)
        bot.edit_message_text(
            "📤 <b>Session clear kora hoyeche!</b>\nNotun kore suru korte apnar <b>.xlsx</b> file-ti upload korun:",
            chat_id,
            call.message.message_id,
            parse_mode="HTML"
        )
        bot.answer_callback_query(call.id)
        return

    session = user_sessions.get(chat_id)
    if not session or not session.get('items'):
        bot.answer_callback_query(call.id, "⚠️ Session expire hoye geche! Abar file upload korun.", show_alert=True)
        return

    items = session['items']
    total = len(items)
    current_index = session['index']

    if call.data == "next_item":
        if current_index < total - 1:
            current_index += 1
        else:
            bot.answer_callback_query(call.id, "✅ Apni file-er ekdom sesh prante chole esechen!", show_alert=True)
            return

    elif call.data == "prev_item":
        if current_index > 0:
            current_index -= 1
        else:
            bot.answer_callback_query(call.id, "⚠️ Etii prothom account!", show_alert=True)
            return

    user_sessions[chat_id]['index'] = current_index
    current_item = items[current_index]

    msg_text = format_account_message(current_index, total, current_item)

    if current_index == total - 1:
        msg_text += "\n\n🎉 <b>এটিই ফাইলের সর্বশেষ অ্যাকাউন্ট!</b>"

    bot.edit_message_text(
        msg_text,
        chat_id,
        call.message.message_id,
        parse_mode="HTML",
        reply_markup=get_navigation_buttons(current_index, total)
    )
    bot.answer_callback_query(call.id, f"Account {current_index + 1}/{total}")

print("🚀 Updated Reader Bot (with Dead UID Filter) is running...")
bot.infinity_polling(timeout=60, long_polling_timeout=30)
