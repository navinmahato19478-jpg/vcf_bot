"""Telegram Contact Utility Bot

Install:
    pip install python-telegram-bot openpyxl

Configuration (recommended):
    Windows PowerShell:
        $env:BOT_TOKEN="123456:ABC..."
        $env:BOT_USERNAME="YourBotUsername"
        $env:ADMIN_IDS="123456789,987654321"

    Linux/macOS:
        export BOT_TOKEN="123456:ABC..."
        export BOT_USERNAME="YourBotUsername"
        export ADMIN_IDS="123456789,987654321"

Run:
    python contact_utility_bot_complete.py
"""

import csv
import math
import os
import re
import shutil
import sqlite3
import uuid
from pathlib import Path

from telegram import (
    Update,
    BotCommand,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

try:
    import openpyxl
except ImportError:
    openpyxl = None


# =========================================================
# CONFIGURATION
# =========================================================

BOT_TOKEN = "9042307549:AAHHi2fWqUAiOjTUCKQ6hpVJPMnX0xhWVlE"
BOT_USERNAME = os.getenv("BOT_USERNAME", "YOUR_BOT_USERNAME").lstrip("@")
ADMIN_IDS = {
    int(value.strip())
    for value in os.getenv("ADMIN_IDS", "").split(",")
    if value.strip().isdigit()
}

BASE_DIR = Path("bot_files")
BASE_DIR.mkdir(parents=True, exist_ok=True)
DATABASE_PATH = BASE_DIR / "bot.db"
MAX_UPLOAD_FILES = 10
MAX_OUTPUT_FILES = 100
MAX_PREVIEW_CHARACTERS = 3500

# All guided operations use one conversation state.
WORKFLOW = 1

FILE_COMMANDS = {
    "to_vcf", "txt_to_csv", "to_txt", "add", "delete", "renamectc",
    "renamefile", "merge", "split", "count", "nodup", "getname",
    "getconten",
}


# =========================================================
# DATABASE
# =========================================================

def db_connection():
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def initialize_database():
    with db_connection() as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                username TEXT,
                full_name TEXT,
                plan TEXT NOT NULL DEFAULT 'Free',
                coins INTEGER NOT NULL DEFAULT 0,
                referred_by INTEGER,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS settings (
                user_id INTEGER PRIMARY KEY,
                default_split TEXT NOT NULL DEFAULT '100',
                number_format TEXT NOT NULL DEFAULT 'Keep original',
                duplicate_handling TEXT NOT NULL DEFAULT 'Keep all',
                language TEXT NOT NULL DEFAULT 'English'
            );
            """
        )


def register_user(user, referred_by=None):
    if not user:
        return

    with db_connection() as connection:
        existing = connection.execute(
            "SELECT user_id FROM users WHERE user_id = ?", (user.id,)
        ).fetchone()

        connection.execute(
            """
            INSERT INTO users (user_id, username, full_name, referred_by)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                username = excluded.username,
                full_name = excluded.full_name
            """,
            (user.id, user.username, user.full_name, referred_by),
        )
        connection.execute(
            "INSERT OR IGNORE INTO settings (user_id) VALUES (?)", (user.id,)
        )

        # Award once, only when the referred user is first registered.
        if not existing and referred_by and referred_by != user.id:
            referrer = connection.execute(
                "SELECT user_id FROM users WHERE user_id = ?", (referred_by,)
            ).fetchone()
            if referrer:
                connection.execute(
                    "UPDATE users SET coins = coins + 10 WHERE user_id = ?",
                    (referred_by,),
                )


def get_user_record(user_id):
    with db_connection() as connection:
        return connection.execute(
            "SELECT * FROM users WHERE user_id = ?", (user_id,)
        ).fetchone()


def get_settings(user_id):
    with db_connection() as connection:
        connection.execute(
            "INSERT OR IGNORE INTO settings (user_id) VALUES (?)", (user_id,)
        )
        return connection.execute(
            "SELECT * FROM settings WHERE user_id = ?", (user_id,)
        ).fetchone()


# =========================================================
# GENERIC HELPERS
# =========================================================

def user_dir(user_id: int) -> Path:
    folder = BASE_DIR / str(user_id)
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def clear_user_temp_files(user_id: int):
    folder = user_dir(user_id)
    for item in folder.iterdir():
        try:
            if item.is_file():
                item.unlink()
            elif item.is_dir():
                shutil.rmtree(item)
        except OSError:
            pass


def safe_filename(value: str, default="contacts") -> str:
    value = re.sub(r'[\\/:*?"<>|]', "_", str(value).strip()).strip(". ")
    return value or default


def clean_phone(value) -> str:
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    plus = text.startswith("+")
    digits = re.sub(r"\D", "", text)
    return f"+{digits}" if plus and digits else digits


def valid_phone(value) -> bool:
    length = len(clean_phone(value).lstrip("+"))
    return 6 <= length <= 16


def parse_txt_contacts(text: str):
    contacts = []
    for index, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue

        if "," in line:
            name, number = line.rsplit(",", 1)
        elif ":" in line:
            name, number = line.rsplit(":", 1)
        else:
            name, number = f"Contact {index}", line

        phone = clean_phone(number)
        if valid_phone(phone):
            contacts.append((name.strip() or f"Contact {index}", phone))
    return contacts


def parse_vcf(text: str):
    contacts = []
    for card in re.split(r"BEGIN:VCARD", text, flags=re.IGNORECASE):
        if not card.strip():
            continue
        name_match = re.search(
            r"^FN(?:;[^:]*)?:(.+)$", card, re.MULTILINE | re.IGNORECASE
        )
        phone_matches = re.findall(
            r"^TEL(?:;[^:]*)?:(.+)$", card, re.MULTILINE | re.IGNORECASE
        )
        name = name_match.group(1).strip() if name_match else "Unknown"
        for number in phone_matches:
            phone = clean_phone(number)
            if valid_phone(phone):
                contacts.append((name, phone))
    return contacts


def contacts_to_vcf(contacts):
    cards = []
    for name, phone in contacts:
        escaped_name = str(name).replace("\n", " ").replace(";", r"\;")
        cards.append(
            "BEGIN:VCARD\n"
            "VERSION:3.0\n"
            f"FN:{escaped_name}\n"
            f"TEL;TYPE=CELL:{phone}\n"
            "END:VCARD"
        )
    return "\n".join(cards)


def read_xlsx_contacts(path: Path):
    if openpyxl is None:
        raise RuntimeError("Install XLSX support with: pip install openpyxl")

    contacts = []
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for worksheet in workbook.worksheets:
            for row_number, row in enumerate(worksheet.iter_rows(values_only=True), 1):
                values = [value for value in row if value is not None]
                if not values:
                    continue

                name = None
                number = None
                for value in values:
                    if valid_phone(value):
                        number = clean_phone(value)
                    elif name is None:
                        name = str(value).strip()

                if number:
                    contacts.append((name or f"Contact {row_number}", number))
    finally:
        workbook.close()
    return contacts


def read_contacts(path: Path):
    extension = path.suffix.lower()
    if extension == ".vcf":
        return parse_vcf(path.read_text(encoding="utf-8", errors="ignore"))
    if extension == ".txt":
        return parse_txt_contacts(path.read_text(encoding="utf-8", errors="ignore"))
    if extension == ".xlsx":
        return read_xlsx_contacts(path)
    raise ValueError(f"Unsupported format: {extension}")


def write_txt(path: Path, contacts, output_format="Name,Number"):
    if output_format == "Number only":
        content = "\n".join(phone for _, phone in contacts)
    elif output_format == "Name only":
        content = "\n".join(name for name, _ in contacts)
    else:
        content = "\n".join(f"{name},{phone}" for name, phone in contacts)
    path.write_text(content, encoding="utf-8")


def write_csv(path: Path, contacts):
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output)
        writer.writerow(["Name", "Phone"])
        writer.writerows(contacts)


def split_chunks(values, method, amount):
    if not values:
        return []
    amount = max(1, int(amount))
    if method == "Number of files":
        chunk_size = math.ceil(len(values) / amount)
    else:
        chunk_size = amount
    return [values[i:i + chunk_size] for i in range(0, len(values), chunk_size)]


def unique_contacts(contacts, mode="Number", keep="First"):
    iterable = contacts if keep == "First" else list(reversed(contacts))
    seen = set()
    result = []

    for name, phone in iterable:
        if mode == "Name":
            key = name.casefold()
        elif mode == "Both":
            key = (name.casefold(), clean_phone(phone))
        else:
            key = clean_phone(phone)
        if key not in seen:
            seen.add(key)
            result.append((name, phone))

    if keep == "Last":
        result.reverse()
    return result


def choice_keyboard(options, prefix="wf"):
    rows = []
    for index in range(0, len(options), 2):
        rows.append([
            InlineKeyboardButton(option, callback_data=f"{prefix}|{option}")
            for option in options[index:index + 2]
        ])
    return InlineKeyboardMarkup(rows)


async def send_path(update: Update, path: Path, caption=None):
    with path.open("rb") as file:
        await update.effective_message.reply_document(
            document=file, filename=path.name, caption=caption
        )


# =========================================================
# WORKFLOW DEFINITIONS
# =========================================================

WORKFLOWS = {
    "to_vcf": {
        "extensions": [".txt", ".xlsx"], "multiple": True,
        "questions": [
            ("prefix", "📝 Enter contact prefix:", "text", None),
            ("filename", "💾 Enter output filename:", "text", None),
            ("split", "🔢 Select contacts per file:", "choice", ["50", "100", "200", "All"]),
        ],
    },
    "txt_to_csv": {
        "extensions": [".txt", ".xlsx"], "multiple": True,
        "questions": [
            ("prefix", "📝 Enter contact prefix:", "text", None),
            ("filename", "💾 Enter output filename:", "text", None),
            ("split", "🔢 Select contacts per file:", "choice", ["50", "100", "200", "All"]),
        ],
    },
    "to_txt": {
        "extensions": [".vcf"], "multiple": True,
        "questions": [
            ("format", "📄 Select output format:", "choice", ["Name,Number", "Number only", "Name only"]),
            ("filename", "💾 Enter output filename:", "text", None),
        ],
    },
    "add": {
        "extensions": [".vcf"], "multiple": False,
        "questions": [
            ("method", "➕ Select add method:", "choice", ["Manual entry", "Upload file"]),
            ("add_data", "📝 Paste contacts as Name,Number, or upload content as text.\nExample: Ravi,+919876543210", "text", None),
            ("filename", "💾 Enter output filename:", "text", None),
        ],
    },
    "delete": {
        "extensions": [".vcf"], "multiple": False,
        "questions": [
            ("method", "🗑 Delete by:", "choice", ["Phone number", "Contact name", "Position"]),
            ("value", "✏️ Enter phone, exact name, position, or range such as 1-50:", "text", None),
            ("filename", "💾 Enter output filename:", "text", None),
        ],
    },
    "renamectc": {
        "extensions": [".vcf"], "multiple": False,
        "questions": [
            ("mode", "✏️ Select rename mode:", "choice", ["Prefix", "Find & Replace", "Sequential"]),
            ("value", "📝 Enter value. For Find & Replace use old|new:", "text", None),
            ("filename", "💾 Enter output filename:", "text", None),
        ],
    },
    "renamefile": {
        "extensions": [".vcf", ".txt", ".xlsx", ".csv"], "multiple": False,
        "questions": [("filename", "💾 Enter the new filename including extension:", "text", None)],
    },
    "merge": {
        "extensions": [".vcf", ".txt"], "multiple": True,
        "questions": [
            ("dedupe", "♻️ Remove duplicates?", "choice", ["Yes", "No"]),
            ("format", "📄 Select output format:", "choice", ["VCF", "TXT"]),
            ("filename", "💾 Enter output filename:", "text", None),
        ],
    },
    "split": {
        "extensions": [".vcf", ".txt"], "multiple": False,
        "questions": [
            ("method", "✂️ Select split method:", "choice", ["Contacts per file", "Number of files"]),
            ("value", "🔢 Enter a positive number:", "number", None),
            ("filename", "💾 Enter base output filename:", "text", None),
        ],
    },
    "count": {
        "extensions": [".vcf", ".txt", ".xlsx"], "multiple": True,
        "questions": [],
    },
    "nodup": {
        "extensions": [".vcf", ".txt"], "multiple": True,
        "questions": [
            ("check", "♻️ Duplicate check:", "choice", ["Number", "Name", "Both"]),
            ("keep", "📌 Keep which record?", "choice", ["First", "Last"]),
            ("filename", "💾 Enter output filename:", "text", None),
        ],
    },
    "getname": {
        "extensions": [".vcf"], "multiple": True,
        "questions": [
            ("format", "📋 Select output:", "choice", ["Names only", "Name + Number", "Unique names"]),
            ("filename", "💾 Enter output filename:", "text", None),
        ],
    },
    "getconten": {
        "extensions": [".txt"], "multiple": True,
        "questions": [
            ("action", "📄 Select action:", "choice", ["Preview", "Extract numbers", "Extract names", "Clean lines"]),
            ("filename", "💾 Enter output filename:", "text", None),
        ],
    },
}


# =========================================================
# START, HELP AND GUIDED FILE WORKFLOWS
# =========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    referred_by = None
    if context.args and context.args[0].isdigit():
        referred_by = int(context.args[0])
    register_user(update.effective_user, referred_by)

    keyboard = ReplyKeyboardMarkup(
        [
            ["/to_vcf", "/txt_to_csv", "/to_txt"],
            ["/manual", "/add", "/delete"],
            ["/renamectc", "/renamefile", "/merge"],
            ["/split", "/count", "/nodup"],
            ["/getname", "/generate", "/getconten"],
            ["/setting", "/status", "/vip"],
            ["/referral", "/admin", "/help"],
        ],
        resize_keyboard=True,
    )
    await update.message.reply_text(
        f"Hello {update.effective_user.first_name}!\n\n"
        "Welcome to Contact Utility Bot. Select a command below.",
        reply_markup=keyboard,
    )


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = """📖 Available commands

/to_vcf - TXT/XLSX → VCF
/txt_to_csv - TXT/XLSX → CSV
/to_txt - VCF → TXT
/manual - Paste contacts and generate VCF
/add - Add contacts to a VCF
/delete - Delete contacts from a VCF
/renamectc - Rename VCF contacts
/renamefile - Rename an uploaded file
/merge - Merge VCF/TXT files
/split - Split VCF/TXT files
/count - Show contact statistics
/nodup - Remove duplicate contacts
/getname - Extract contact names
/generate - Generate sequential contacts
/getconten - Process TXT content
/setting - User settings
/status - Account status
/vip - VIP plans
/referral - Referral information
/admin - Admin menu

During an upload workflow, use /done to continue or /cancel to stop."""
    await update.message.reply_text(text)


async def begin_file_workflow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    command = update.message.text.split()[0].split("@")[0].lstrip("/")
    definition = WORKFLOWS[command]
    context.user_data.clear()
    context.user_data.update({
        "workflow": command,
        "files": [],
        "answers": {},
        "question_index": 0,
        "stage": "upload",
    })
    clear_user_temp_files(update.effective_user.id)
    allowed = ", ".join(definition["extensions"])
    multiple = "one or more" if definition["multiple"] else "one"
    await update.message.reply_text(
        f"📥 Upload {multiple} file(s): {allowed}\n\n"
        "After uploading, type /done.\nType /cancel to stop."
    )
    return WORKFLOW


async def receive_workflow_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    workflow = context.user_data.get("workflow")
    if not workflow or context.user_data.get("stage") != "upload":
        await update.message.reply_text("Start a command before uploading a file.")
        return WORKFLOW

    definition = WORKFLOWS[workflow]
    document = update.message.document
    original_name = document.file_name or "file"
    extension = Path(original_name).suffix.lower()

    if extension not in definition["extensions"]:
        await update.message.reply_text(
            f"❌ Unsupported file. Allowed: {', '.join(definition['extensions'])}"
        )
        return WORKFLOW

    files = context.user_data["files"]
    if not definition["multiple"] and files:
        await update.message.reply_text("❌ This command accepts only one file. Type /done.")
        return WORKFLOW
    if len(files) >= MAX_UPLOAD_FILES:
        await update.message.reply_text(f"❌ Maximum {MAX_UPLOAD_FILES} files allowed.")
        return WORKFLOW

    folder = user_dir(update.effective_user.id)
    path = folder / f"{uuid.uuid4().hex}_{safe_filename(Path(original_name).stem)}{extension}"
    telegram_file = await document.get_file()
    await telegram_file.download_to_drive(custom_path=path)
    files.append(path)

    await update.message.reply_text(
        f"✅ File received: {original_name}\nFiles uploaded: {len(files)}\n\n"
        "Upload another file or type /done."
    )
    return WORKFLOW


async def workflow_done(update: Update, context: ContextTypes.DEFAULT_TYPE):
    workflow = context.user_data.get("workflow")
    if not workflow:
        await update.message.reply_text("No active operation.")
        return ConversationHandler.END
    if context.user_data.get("stage") != "upload":
        await update.message.reply_text("Please answer the current question.")
        return WORKFLOW
    if not context.user_data.get("files"):
        await update.message.reply_text("❌ Upload at least one file first.")
        return WORKFLOW

    context.user_data["stage"] = "questions"
    context.user_data["question_index"] = 0
    return await ask_next_question(update, context)


async def ask_next_question(update: Update, context: ContextTypes.DEFAULT_TYPE):
    workflow = context.user_data["workflow"]
    questions = WORKFLOWS[workflow]["questions"]
    index = context.user_data["question_index"]

    if index >= len(questions):
        return await process_workflow(update, context)

    key, prompt, question_type, options = questions[index]
    context.user_data["current_question"] = (key, question_type)
    markup = choice_keyboard(options) if question_type == "choice" else None
    await update.effective_message.reply_text(prompt, reply_markup=markup)
    return WORKFLOW


async def receive_workflow_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if context.user_data.get("stage") == "upload":
        await update.message.reply_text("Upload a supported file or type /done.")
        return WORKFLOW

    current = context.user_data.get("current_question")
    if not current:
        return WORKFLOW
    key, question_type = current
    value = update.message.text.strip()

    if question_type == "choice":
        question = WORKFLOWS[context.user_data["workflow"]]["questions"][context.user_data["question_index"]]
        options = question[3]
        matched = next((option for option in options if option.casefold() == value.casefold()), None)
        if not matched:
            await update.message.reply_text(f"❌ Select one of: {', '.join(options)}")
            return WORKFLOW
        value = matched
    elif question_type == "number":
        if not value.isdigit() or int(value) <= 0:
            await update.message.reply_text("❌ Enter a positive whole number.")
            return WORKFLOW
        value = int(value)
    elif not value:
        await update.message.reply_text("❌ Value cannot be empty.")
        return WORKFLOW

    context.user_data["answers"][key] = value
    context.user_data["question_index"] += 1
    return await ask_next_question(update, context)


async def receive_workflow_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not context.user_data.get("current_question"):
        await query.edit_message_text("This operation has expired. Start again.")
        return ConversationHandler.END

    _, value = query.data.split("|", 1)
    key, _ = context.user_data["current_question"]
    context.user_data["answers"][key] = value
    context.user_data["question_index"] += 1
    await query.edit_message_text(f"✅ Selected: {value}")
    return await ask_next_question(update, context)


async def cancel_workflow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_user_temp_files(update.effective_user.id)
    context.user_data.clear()
    await update.message.reply_text("❌ Operation cancelled.")
    return ConversationHandler.END


# =========================================================
# PROCESSORS
# =========================================================

async def process_workflow(update: Update, context: ContextTypes.DEFAULT_TYPE):
    workflow = context.user_data["workflow"]
    files = context.user_data["files"]
    answers = context.user_data["answers"]
    message = update.effective_message
    await message.reply_text("⏳ Processing...")

    try:
        if workflow in {"to_vcf", "txt_to_csv"}:
            await process_convert_number_files(update, files, answers, workflow)
        elif workflow == "to_txt":
            await process_to_txt(update, files, answers)
        elif workflow == "add":
            await process_add(update, files[0], answers)
        elif workflow == "delete":
            await process_delete(update, files[0], answers)
        elif workflow == "renamectc":
            await process_rename_contacts(update, files[0], answers)
        elif workflow == "renamefile":
            await process_rename_file(update, files[0], answers)
        elif workflow == "merge":
            await process_merge(update, files, answers)
        elif workflow == "split":
            await process_split(update, files[0], answers)
        elif workflow == "count":
            await process_count(update, files)
        elif workflow == "nodup":
            await process_nodup(update, files, answers)
        elif workflow == "getname":
            await process_getname(update, files, answers)
        elif workflow == "getconten":
            await process_getcontent(update, files, answers)
    except Exception as error:
        await message.reply_text(f"❌ Processing failed:\n{error}")
    finally:
        context.user_data.clear()

    return ConversationHandler.END


async def process_convert_number_files(update, files, answers, workflow):
    source_contacts = []
    for path in files:
        source_contacts.extend(read_contacts(path))

    if not source_contacts:
        raise ValueError("No valid phone numbers found.")

    prefix = answers["prefix"].strip()
    filename = safe_filename(answers["filename"])
    contacts = [(f"{prefix} {index}", phone) for index, (_, phone) in enumerate(source_contacts, 1)]
    size = len(contacts) if answers["split"] == "All" else int(answers["split"])
    chunks = split_chunks(contacts, "Contacts per file", size)
    if len(chunks) > MAX_OUTPUT_FILES:
        raise ValueError(f"Output exceeds the {MAX_OUTPUT_FILES}-file limit.")

    extension = ".vcf" if workflow == "to_vcf" else ".csv"
    folder = user_dir(update.effective_user.id)
    for index, chunk in enumerate(chunks, 1):
        suffix = f" {index}" if len(chunks) > 1 else ""
        output = folder / f"{filename}{suffix}{extension}"
        if workflow == "to_vcf":
            output.write_text(contacts_to_vcf(chunk), encoding="utf-8")
        else:
            write_csv(output, chunk)
        await send_path(update, output)

    await update.effective_message.reply_text(
        f"✅ Completed. Contacts: {len(contacts)} | Files: {len(chunks)}"
    )


async def process_to_txt(update, files, answers):
    contacts = []
    for path in files:
        contacts.extend(read_contacts(path))
    filename = safe_filename(answers["filename"])
    output = user_dir(update.effective_user.id) / f"{filename}.txt"
    write_txt(output, contacts, answers["format"])
    await send_path(update, output, f"✅ Extracted {len(contacts)} contacts.")


async def process_add(update, path, answers):
    contacts = read_contacts(path)
    # Both options accept pasted Name,Number content in this version.
    additions = parse_txt_contacts(answers["add_data"])
    if not additions:
        raise ValueError("No valid contacts were supplied to add.")
    contacts.extend(additions)
    filename = safe_filename(answers["filename"])
    output = user_dir(update.effective_user.id) / f"{filename}.vcf"
    output.write_text(contacts_to_vcf(contacts), encoding="utf-8")
    await send_path(update, output, f"✅ Added {len(additions)} contact(s).")


async def process_delete(update, path, answers):
    contacts = read_contacts(path)
    method, value = answers["method"], answers["value"].strip()
    original_count = len(contacts)

    if method == "Phone number":
        target = clean_phone(value)
        contacts = [(n, p) for n, p in contacts if clean_phone(p) != target]
    elif method == "Contact name":
        contacts = [(n, p) for n, p in contacts if n.casefold() != value.casefold()]
    else:
        positions = set()
        if "-" in value:
            start, end = value.split("-", 1)
            if not start.strip().isdigit() or not end.strip().isdigit():
                raise ValueError("Position range must look like 1-50.")
            positions.update(range(int(start), int(end) + 1))
        elif value.isdigit():
            positions.add(int(value))
        else:
            raise ValueError("Position must be a number or range such as 1-50.")
        contacts = [contact for index, contact in enumerate(contacts, 1) if index not in positions]

    filename = safe_filename(answers["filename"])
    output = user_dir(update.effective_user.id) / f"{filename}.vcf"
    output.write_text(contacts_to_vcf(contacts), encoding="utf-8")
    await send_path(update, output, f"✅ Deleted {original_count - len(contacts)} contact(s).")


async def process_rename_contacts(update, path, answers):
    contacts = read_contacts(path)
    mode, value = answers["mode"], answers["value"]
    renamed = []

    if mode == "Find & Replace":
        if "|" not in value:
            raise ValueError("Use old|new for Find & Replace.")
        old, new = value.split("|", 1)
        renamed = [(name.replace(old, new), phone) for name, phone in contacts]
    elif mode in {"Prefix", "Sequential"}:
        renamed = [(f"{value} {index}", phone) for index, (_, phone) in enumerate(contacts, 1)]

    filename = safe_filename(answers["filename"])
    output = user_dir(update.effective_user.id) / f"{filename}.vcf"
    output.write_text(contacts_to_vcf(renamed), encoding="utf-8")
    await send_path(update, output, f"✅ Renamed {len(renamed)} contacts.")


async def process_rename_file(update, path, answers):
    new_name = safe_filename(answers["filename"])
    if not Path(new_name).suffix:
        new_name += path.suffix
    output = user_dir(update.effective_user.id) / new_name
    shutil.copy2(path, output)
    await send_path(update, output, "✅ File renamed.")


async def process_merge(update, files, answers):
    contacts = []
    for path in files:
        contacts.extend(read_contacts(path))
    if answers["dedupe"] == "Yes":
        contacts = unique_contacts(contacts, "Number", "First")

    filename = safe_filename(answers["filename"])
    extension = ".vcf" if answers["format"] == "VCF" else ".txt"
    output = user_dir(update.effective_user.id) / f"{filename}{extension}"
    if extension == ".vcf":
        output.write_text(contacts_to_vcf(contacts), encoding="utf-8")
    else:
        write_txt(output, contacts)
    await send_path(update, output, f"✅ Merged {len(files)} files and {len(contacts)} contacts.")


async def process_split(update, path, answers):
    contacts = read_contacts(path)
    chunks = split_chunks(contacts, answers["method"], answers["value"])
    if len(chunks) > MAX_OUTPUT_FILES:
        raise ValueError(f"Output exceeds the {MAX_OUTPUT_FILES}-file limit.")

    filename = safe_filename(answers["filename"])
    extension = path.suffix.lower()
    folder = user_dir(update.effective_user.id)
    for index, chunk in enumerate(chunks, 1):
        output = folder / f"{filename} {index}{extension}"
        if extension == ".vcf":
            output.write_text(contacts_to_vcf(chunk), encoding="utf-8")
        else:
            write_txt(output, chunk)
        await send_path(update, output)
    await update.effective_message.reply_text(f"✅ Created {len(chunks)} files.")


async def process_count(update, files):
    contacts = []
    for path in files:
        contacts.extend(read_contacts(path))
    numbers = [clean_phone(phone) for _, phone in contacts]
    unique_numbers = len(set(numbers))
    named = sum(1 for name, _ in contacts if name and not name.startswith("Contact ") and name != "Unknown")
    await update.effective_message.reply_text(
        "📊 Contact statistics\n\n"
        f"Files: {len(files)}\n"
        f"Total contacts: {len(contacts)}\n"
        f"Unique numbers: {unique_numbers}\n"
        f"Duplicate entries: {len(contacts) - unique_numbers}\n"
        f"Named contacts: {named}\n"
        f"Unnamed/generated names: {len(contacts) - named}"
    )


async def process_nodup(update, files, answers):
    contacts = []
    for path in files:
        contacts.extend(read_contacts(path))
    unique = unique_contacts(contacts, answers["check"], answers["keep"])
    filename = safe_filename(answers["filename"])
    extension = files[0].suffix.lower()
    output = user_dir(update.effective_user.id) / f"{filename}{extension}"
    if extension == ".vcf":
        output.write_text(contacts_to_vcf(unique), encoding="utf-8")
    else:
        write_txt(output, unique)
    await send_path(update, output, f"✅ Removed {len(contacts) - len(unique)} duplicate(s).")


async def process_getname(update, files, answers):
    contacts = []
    for path in files:
        contacts.extend(read_contacts(path))
    selected = answers["format"]
    if selected == "Unique names":
        lines = list(dict.fromkeys(name for name, _ in contacts))
    elif selected == "Name + Number":
        lines = [f"{name},{phone}" for name, phone in contacts]
    else:
        lines = [name for name, _ in contacts]
    output = user_dir(update.effective_user.id) / f"{safe_filename(answers['filename'])}.txt"
    output.write_text("\n".join(lines), encoding="utf-8")
    await send_path(update, output, f"✅ Extracted {len(lines)} entries.")


async def process_getcontent(update, files, answers):
    text = "\n".join(path.read_text(encoding="utf-8", errors="ignore") for path in files)
    action = answers["action"]

    if action == "Preview":
        preview = text[:MAX_PREVIEW_CHARACTERS]
        if len(text) > MAX_PREVIEW_CHARACTERS:
            preview += "\n\n[Preview truncated]"
        await update.effective_message.reply_text(preview or "The file is empty.")
        return

    lines = []
    if action == "Extract numbers":
        lines = [phone for _, phone in parse_txt_contacts(text)]
    elif action == "Extract names":
        lines = [name for name, _ in parse_txt_contacts(text)]
    elif action == "Clean lines":
        lines = [line.strip() for line in text.splitlines() if line.strip()]

    output = user_dir(update.effective_user.id) / f"{safe_filename(answers['filename'])}.txt"
    output.write_text("\n".join(lines), encoding="utf-8")
    await send_path(update, output, f"✅ Created {len(lines)} lines.")


# =========================================================
# MANUAL AND GENERATE CONVERSATIONS
# =========================================================

MANUAL_CONTACTS, MANUAL_PREFIX, MANUAL_FILENAME, MANUAL_SPLIT = range(10, 14)
GENERATE_PREFIX, GENERATE_QTY, GENERATE_START, GENERATE_FILENAME, GENERATE_SPLIT = range(20, 25)


async def manual_start(update, context):
    context.user_data.clear()
    await update.message.reply_text(
        "📝 Paste contacts, one per line. Supported formats:\n"
        "Ravi,+919876543210\n+919876543211\n\nThen type /done."
    )
    context.user_data["manual_lines"] = []
    return MANUAL_CONTACTS


async def manual_receive(update, context):
    context.user_data["manual_lines"].append(update.message.text)
    await update.message.reply_text("✅ Contacts received. Send more or type /done.")
    return MANUAL_CONTACTS


async def manual_done(update, context):
    if not context.user_data.get("manual_lines"):
        await update.message.reply_text("❌ Paste contacts first.")
        return MANUAL_CONTACTS
    await update.message.reply_text("📝 Enter contact prefix:")
    return MANUAL_PREFIX


async def manual_prefix(update, context):
    context.user_data["prefix"] = update.message.text.strip()
    await update.message.reply_text("💾 Enter output filename:")
    return MANUAL_FILENAME


async def manual_filename(update, context):
    context.user_data["filename"] = safe_filename(update.message.text)
    await update.message.reply_text(
        "🔢 Select contacts per file:",
        reply_markup=choice_keyboard(["50", "100", "200", "All"], "manual"),
    )
    return MANUAL_SPLIT


async def manual_split(update, context):
    query = update.callback_query
    await query.answer()
    _, selected = query.data.split("|", 1)
    source = parse_txt_contacts("\n".join(context.user_data["manual_lines"]))
    prefix = context.user_data["prefix"]
    contacts = [(f"{prefix} {i}", phone) for i, (_, phone) in enumerate(source, 1)]
    size = len(contacts) if selected == "All" else int(selected)
    chunks = split_chunks(contacts, "Contacts per file", size)
    folder = user_dir(update.effective_user.id)
    for index, chunk in enumerate(chunks, 1):
        suffix = f" {index}" if len(chunks) > 1 else ""
        output = folder / f"{context.user_data['filename']}{suffix}.vcf"
        output.write_text(contacts_to_vcf(chunk), encoding="utf-8")
        await send_path(update, output)
    await query.edit_message_text(f"✅ Created {len(chunks)} file(s).")
    context.user_data.clear()
    return ConversationHandler.END


async def generate_start(update, context):
    context.user_data.clear()
    await update.message.reply_text("📝 Enter contact prefix:")
    return GENERATE_PREFIX


async def generate_prefix(update, context):
    context.user_data["prefix"] = update.message.text.strip()
    await update.message.reply_text("🔢 Enter quantity:")
    return GENERATE_QTY


async def generate_qty(update, context):
    if not update.message.text.isdigit() or int(update.message.text) <= 0:
        await update.message.reply_text("❌ Enter a positive whole number.")
        return GENERATE_QTY
    context.user_data["quantity"] = int(update.message.text)
    await update.message.reply_text(
        "☎️ Enter starting phone number. Numbers will increase sequentially:"
    )
    return GENERATE_START


async def generate_start_number(update, context):
    phone = clean_phone(update.message.text)
    if not valid_phone(phone):
        await update.message.reply_text("❌ Enter a valid starting phone number.")
        return GENERATE_START
    context.user_data["start_number"] = phone
    await update.message.reply_text("💾 Enter output filename:")
    return GENERATE_FILENAME


async def generate_filename(update, context):
    context.user_data["filename"] = safe_filename(update.message.text)
    await update.message.reply_text(
        "🔢 Select contacts per file:",
        reply_markup=choice_keyboard(["50", "100", "200", "All"], "generate"),
    )
    return GENERATE_SPLIT


async def generate_split(update, context):
    query = update.callback_query
    await query.answer()
    _, selected = query.data.split("|", 1)
    quantity = context.user_data["quantity"]
    original = context.user_data["start_number"]
    plus = original.startswith("+")
    start_number = int(original.lstrip("+"))
    width = len(original.lstrip("+"))
    contacts = []
    for index in range(quantity):
        number = str(start_number + index).zfill(width)
        if plus:
            number = "+" + number
        contacts.append((f"{context.user_data['prefix']} {index + 1}", number))

    size = quantity if selected == "All" else int(selected)
    chunks = split_chunks(contacts, "Contacts per file", size)
    folder = user_dir(update.effective_user.id)
    for index, chunk in enumerate(chunks, 1):
        suffix = f" {index}" if len(chunks) > 1 else ""
        output = folder / f"{context.user_data['filename']}{suffix}.vcf"
        output.write_text(contacts_to_vcf(chunk), encoding="utf-8")
        await send_path(update, output)
    await query.edit_message_text(f"✅ Generated {quantity} contacts in {len(chunks)} file(s).")
    context.user_data.clear()
    return ConversationHandler.END


# =========================================================
# SETTINGS, STATUS, VIP, REFERRAL AND ADMIN
# =========================================================

async def settings_command(update, context):
    settings = get_settings(update.effective_user.id)
    keyboard = choice_keyboard(
        ["Default split", "Number format", "Duplicate handling", "Language"],
        "setting",
    )
    await update.message.reply_text(
        "⚙️ Settings\n\n"
        f"Default split: {settings['default_split']}\n"
        f"Number format: {settings['number_format']}\n"
        f"Duplicate handling: {settings['duplicate_handling']}\n"
        f"Language: {settings['language']}",
        reply_markup=keyboard,
    )


async def settings_button(update, context):
    query = update.callback_query
    await query.answer()
    _, setting = query.data.split("|", 1)
    options = {
        "Default split": ["50", "100", "200", "All"],
        "Number format": ["Keep original", "Digits only", "Leading plus"],
        "Duplicate handling": ["Keep all", "Remove by number"],
        "Language": ["English"],
    }[setting]
    context.user_data["setting_name"] = setting
    await query.edit_message_text(
        f"Select {setting}:", reply_markup=choice_keyboard(options, "setvalue")
    )


async def setting_value_button(update, context):
    query = update.callback_query
    await query.answer()
    _, value = query.data.split("|", 1)
    setting = context.user_data.pop("setting_name", None)
    columns = {
        "Default split": "default_split",
        "Number format": "number_format",
        "Duplicate handling": "duplicate_handling",
        "Language": "language",
    }
    if not setting:
        await query.edit_message_text("Setting session expired.")
        return
    column = columns[setting]
    with db_connection() as connection:
        connection.execute(
            f"UPDATE settings SET {column} = ? WHERE user_id = ?",
            (value, update.effective_user.id),
        )
    await query.edit_message_text(f"✅ {setting} updated to: {value}")


async def status_command(update, context):
    register_user(update.effective_user)
    record = get_user_record(update.effective_user.id)
    with db_connection() as connection:
        referrals = connection.execute(
            "SELECT COUNT(*) AS total FROM users WHERE referred_by = ?",
            (update.effective_user.id,),
        ).fetchone()["total"]
    await update.message.reply_text(
        "📊 Account status\n\n"
        f"User ID: {record['user_id']}\n"
        f"Name: {record['full_name']}\n"
        f"Plan: {record['plan']}\n"
        f"Coins: {record['coins']}\n"
        f"Referrals: {referrals}\n"
        "Status: Active"
    )


async def vip_command(update, context):
    await update.message.reply_text(
        "💎 Choose a VIP plan:\n\n"
        "Monthly: 30 days\nQuarterly: 90 days\nYearly: 365 days\n\n"
        "Payment and activation are not automated. An administrator must approve activation.",
        reply_markup=choice_keyboard(["Monthly", "Quarterly", "Yearly"], "vip"),
    )


async def vip_button(update, context):
    query = update.callback_query
    await query.answer()
    _, plan = query.data.split("|", 1)
    await query.edit_message_text(
        f"Selected plan: {plan}\n\nContact the bot administrator for payment and activation."
    )


async def referral_command(update, context):
    register_user(update.effective_user)
    record = get_user_record(update.effective_user.id)
    with db_connection() as connection:
        referrals = connection.execute(
            "SELECT COUNT(*) AS total FROM users WHERE referred_by = ?",
            (update.effective_user.id,),
        ).fetchone()["total"]
    await update.message.reply_text(
        "🎁 Referral information\n\n"
        f"Link: https://t.me/{BOT_USERNAME}?start={update.effective_user.id}\n"
        f"Successful referrals: {referrals}\n"
        f"Coins: {record['coins']}\n\n"
        "Each new valid referral awards 10 coins."
    )


async def admin_command(update, context):
    if update.effective_user.id not in ADMIN_IDS:
        await update.message.reply_text("⛔ Admin access denied.")
        return
    await update.message.reply_text(
        "🛡 Admin menu:",
        reply_markup=choice_keyboard(["Users", "VIP", "Broadcast", "Statistics"], "admin"),
    )


async def admin_button(update, context):
    query = update.callback_query
    await query.answer()
    if update.effective_user.id not in ADMIN_IDS:
        await query.edit_message_text("⛔ Admin access denied.")
        return
    _, action = query.data.split("|", 1)
    with db_connection() as connection:
        users = connection.execute("SELECT COUNT(*) AS total FROM users").fetchone()["total"]
        vip_users = connection.execute(
            "SELECT COUNT(*) AS total FROM users WHERE plan != 'Free'"
        ).fetchone()["total"]
    if action in {"Users", "Statistics"}:
        text = f"Users: {users}\nVIP users: {vip_users}\nFree users: {users - vip_users}"
    elif action == "VIP":
        text = "VIP management is database-backed. Update the users.plan field after payment approval."
    else:
        text = "Broadcast safety: this build does not send mass messages automatically. Add a confirmation step before enabling broadcasts."
    await query.edit_message_text(text)


# =========================================================
# COMMAND MENU AND MAIN
# =========================================================

async def setup_commands(application):
    commands = [
        BotCommand("start", "Start bot"),
        BotCommand("to_vcf", "Convert TXT/XLSX to VCF"),
        BotCommand("txt_to_csv", "Convert TXT/XLSX to CSV"),
        BotCommand("to_txt", "Convert VCF to TXT"),
        BotCommand("manual", "Manual contact input"),
        BotCommand("add", "Add contacts to VCF"),
        BotCommand("delete", "Delete contacts"),
        BotCommand("renamectc", "Rename contacts"),
        BotCommand("renamefile", "Rename file"),
        BotCommand("merge", "Merge files"),
        BotCommand("split", "Split file"),
        BotCommand("count", "Contact statistics"),
        BotCommand("nodup", "Remove duplicates"),
        BotCommand("getname", "Extract names"),
        BotCommand("generate", "Generate contacts"),
        BotCommand("getconten", "Process TXT content"),
        BotCommand("setting", "Settings"),
        BotCommand("status", "Account status"),
        BotCommand("vip", "VIP plans"),
        BotCommand("referral", "Referral details"),
        BotCommand("admin", "Admin menu"),
        BotCommand("done", "Finish current input"),
        BotCommand("cancel", "Cancel operation"),
        BotCommand("help", "Show commands"),
    ]
    await application.bot.set_my_commands(commands)


def main():
    if BOT_TOKEN == "YOUR_TELEGRAM_BOT_TOKEN":
        raise ValueError("Set BOT_TOKEN before running the bot.")

    initialize_database()
    application = (
        Application.builder().token(BOT_TOKEN).post_init(setup_commands).build()
    )

    file_workflow = ConversationHandler(
        entry_points=[CommandHandler(list(FILE_COMMANDS), begin_file_workflow)],
        states={
            WORKFLOW: [
                CommandHandler("done", workflow_done),
                MessageHandler(filters.Document.ALL, receive_workflow_file),
                CallbackQueryHandler(receive_workflow_button, pattern=r"^wf\|"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_workflow_text),
            ]
        },
        fallbacks=[CommandHandler("cancel", cancel_workflow)],
        allow_reentry=True,
    )

    manual_conversation = ConversationHandler(
        entry_points=[CommandHandler("manual", manual_start)],
        states={
            MANUAL_CONTACTS: [
                CommandHandler("done", manual_done),
                MessageHandler(filters.TEXT & ~filters.COMMAND, manual_receive),
            ],
            MANUAL_PREFIX: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_prefix)],
            MANUAL_FILENAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, manual_filename)],
            MANUAL_SPLIT: [CallbackQueryHandler(manual_split, pattern=r"^manual\|")],
        },
        fallbacks=[CommandHandler("cancel", cancel_workflow)],
        allow_reentry=True,
    )

    generate_conversation = ConversationHandler(
        entry_points=[CommandHandler("generate", generate_start)],
        states={
            GENERATE_PREFIX: [MessageHandler(filters.TEXT & ~filters.COMMAND, generate_prefix)],
            GENERATE_QTY: [MessageHandler(filters.TEXT & ~filters.COMMAND, generate_qty)],
            GENERATE_START: [MessageHandler(filters.TEXT & ~filters.COMMAND, generate_start_number)],
            GENERATE_FILENAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, generate_filename)],
            GENERATE_SPLIT: [CallbackQueryHandler(generate_split, pattern=r"^generate\|")],
        },
        fallbacks=[CommandHandler("cancel", cancel_workflow)],
        allow_reentry=True,
    )

    application.add_handler(file_workflow)
    application.add_handler(manual_conversation)
    application.add_handler(generate_conversation)

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("setting", settings_command))
    application.add_handler(CommandHandler("status", status_command))
    application.add_handler(CommandHandler("vip", vip_command))
    application.add_handler(CommandHandler("referral", referral_command))
    application.add_handler(CommandHandler("admin", admin_command))

    application.add_handler(CallbackQueryHandler(settings_button, pattern=r"^setting\|"))
    application.add_handler(CallbackQueryHandler(setting_value_button, pattern=r"^setvalue\|"))
    application.add_handler(CallbackQueryHandler(vip_button, pattern=r"^vip\|"))
    application.add_handler(CallbackQueryHandler(admin_button, pattern=r"^admin\|"))

    print("Telegram Contact Utility Bot is running...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
