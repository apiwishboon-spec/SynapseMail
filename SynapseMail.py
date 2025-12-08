#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto Mail Center — CustomTkinter Full App (B - Full-featured)
Author: Apiwish Anutarvanichkul (Boon)
Version: 4.1.0 - Full feature set + safe logout + thread/event management + phone input
Requirements: customtkinter
Run: python auto_mail_ctk_full.py
"""

import threading
import imaplib
import smtplib
import email
import email.utils
import webbrowser
import re
from datetime import datetime
import time
import sys

import tkinter as tk
import customtkinter as ctk
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter import messagebox

# -----------------------
# Appearance
# -----------------------
ctk.set_appearance_mode("dark")     # "dark", "light", "system"
ctk.set_default_color_theme("dark-blue")  # "blue", "green", "dark-blue"

# -----------------------
# Globals & Defaults
# -----------------------
EMAIL_ADDRESS = None
EMAIL_PASSWORD = None

DEFAULT_CHECK_INTERVAL = 60  # seconds

GREETING_TEMPLATES = {
    "Friendly 🌈": "Hi {name}! 🎉\nJust wanted to drop in and say hello! Hope you're having an amazing day.\n\n",
    "Professional 📄": "Greetings {name},\nThank you for contacting us. We appreciate your time.\n\n",
    "Tech Nerd 🤖": "[SYSTEM ONLINE] Greetings, {name} 🤖\n$ ssh connection@established\nQuantum entanglement confirmed. Handshake protocol: SUCCESS ✓\n\n",
    "Casual ☕": "What's up {name}? ☕\nJust checking in! Hope everything's going well on your end.\n\n",
    "Enthusiastic 🚀": "HELLO {name}!! 🚀\nSuper excited to connect with you! Let's make something awesome happen!\n\n",
    "Funny 😄": "Yo {name}! 😄\n*Dramatically enters inbox* Hello there! Just sliding into your emails like a pro.\n\n",
    "AI Assistant 🤖": "[AI] Hello {name}! 🤖\n*beep boop* Human detected! My neural networks are pleased to make your acquaintance.\n\n",
    "Sci-Fi Commander 🛸": "Commander {name}, 🛸\n[INCOMING TRANSMISSION]\nThis is Starship Alpha-7. We've detected your signal.\n\n",
}

# Template for auto-reply HTML; {phone} will be replaced with user phone
AUTO_REPLY_HTML_TEMPLATE = """\
<!DOCTYPE html>
<html>
  <body style="margin:0; padding:0; background: linear-gradient(to bottom right, #111827, #0f172a); font-family: Inter, Arial, sans-serif; color:#f8fafc">
    <div style="max-width:560px; margin:60px auto; background:#0b1220; padding:20px; border-radius:12px; box-shadow: 0 8px 30px rgba(0,0,0,0.6);">
      <h3 style="color:#60a5fa; margin-top:0;">Thanks for reaching out 👋</h3>
      <p>We received your message. This is an automatic reply to confirm receipt. We'll get back to you as soon as possible.</p>
      <p style="font-size:13px; color:#9ca3af;">If it's urgent, please call: <strong>{phone}</strong></p>
      <hr style="border:none; border-top:1px solid rgba(255,255,255,0.04); margin:10px 0;">
      <p style="font-size:12px; color:#94a3b8;">Automated system message — no action required.</p>
    </div>
  </body>
</html>
"""

# -----------------------
# Helpers
# -----------------------
def is_valid_email(addr: str) -> bool:
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, addr) is not None

def normalize_phone_digits(text: str) -> str:
    return re.sub(r'\D', '', (text or ""))[:10]  # keep up to 10 digits

def format_phone_dashed(digits: str) -> str:
    d = digits
    if not d:
        return ""
    if len(d) <= 3:
        return d
    if len(d) <= 6:
        return f"{d[:3]}-{d[3:]}"
    return f"{d[:3]}-{d[3:6]}-{d[6:10]}"

# -----------------------
# CTK Login Dialog
# -----------------------
class CTKLoginDialog:
    def __init__(self, parent):
        self.parent = parent
        self.result = None
        self.top = ctk.CTkToplevel(parent)
        self.top.title("🔐 Login")
        self.top.geometry("480x520")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()

        # center
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - (480 // 2)
        y = (self.top.winfo_screenheight() // 2) - (520 // 2)
        self.top.geometry(f"+{x}+{y}")

        self._build_ui()
        self.top.bind("<Return>", lambda e: self._on_submit())
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _build_ui(self):
        header = ctk.CTkLabel(self.top, text="Auto Mail Center — Login", font=ctk.CTkFont(size=20, weight="bold"))
        header.pack(pady=(18, 6))

        desc = ctk.CTkLabel(self.top, text="Use Gmail + App Password (recommended)", font=ctk.CTkFont(size=12))
        desc.pack(pady=(0, 12))

        card = ctk.CTkFrame(self.top, corner_radius=12)
        card.pack(padx=20, pady=6, fill="both", expand=False)

        lbl_email = ctk.CTkLabel(card, text="Email Address", font=ctk.CTkFont(size=11))
        lbl_email.pack(anchor="w", padx=16, pady=(12, 0))
        self.entry_email = ctk.CTkEntry(card, width=420, placeholder_text="you@gmail.com")
        self.entry_email.pack(padx=16, pady=(6, 8))
        self.entry_email.focus()

        lbl_pw = ctk.CTkLabel(card, text="App Password", font=ctk.CTkFont(size=11))
        lbl_pw.pack(anchor="w", padx=16, pady=(8, 0))
        self.entry_pw = ctk.CTkEntry(card, width=420, show="●", placeholder_text="16-char app password")
        self.entry_pw.pack(padx=16, pady=(6, 6))

        # Phone input
        lbl_phone = ctk.CTkLabel(card, text="Phone Number (000-000-0000)", font=ctk.CTkFont(size=11))
        lbl_phone.pack(anchor="w", padx=16, pady=(8, 0))
        self.entry_phone = ctk.CTkEntry(card, width=420, placeholder_text="000-000-0000")
        self.entry_phone.pack(padx=16, pady=(6, 8))

        # Bind a key release to auto-format phone
        def _on_phone_key(event=None):
            cur = self.entry_phone.get()
            digits = normalize_phone_digits(cur)
            formatted = format_phone_dashed(digits)
            # avoid moving cursor to end if already same
            if cur != formatted:
                # set and keep cursor at end (simple approach)
                self.entry_phone.delete(0, tk.END)
                self.entry_phone.insert(0, formatted)

        self.entry_phone.bind("<KeyRelease>", _on_phone_key)

        self.show_pw_var = tk.BooleanVar(value=False)
        chk = ctk.CTkCheckBox(card, text="Show password", variable=self.show_pw_var, command=self._toggle_pw)
        chk.pack(anchor="w", padx=12, pady=(2, 8))

        help_link = ctk.CTkLabel(card, text="Need help with App Passwords?", cursor="hand2", text_color="#60a5fa", font=ctk.CTkFont(size=10))
        help_link.pack(padx=16, pady=(0, 12), anchor="w")
        help_link.bind("<Button-1>", lambda e: webbrowser.open("https://support.google.com/accounts/answer/185833"))

        # Buttons
        btn_frame = ctk.CTkFrame(self.top, fg_color="transparent")
        btn_frame.pack(pady=(10, 18))

        btn_login = ctk.CTkButton(btn_frame, text="Login", width=160, command=self._on_submit)
        btn_login.grid(row=0, column=0, padx=(8, 12))
        btn_cancel = ctk.CTkButton(btn_frame, text="Cancel", width=120, fg_color="#ef4444", hover_color="#f87171", command=self._on_cancel)
        btn_cancel.grid(row=0, column=1, padx=(0, 8))

        self.status_label = ctk.CTkLabel(self.top, text="", font=ctk.CTkFont(size=10))
        self.status_label.pack(pady=(0, 4))

    def _toggle_pw(self):
        self.entry_pw.configure(show="" if self.show_pw_var.get() else "●")

    def _on_submit(self):
        email_val = self.entry_email.get().strip()
        pw_val = self.entry_pw.get().strip()
        phone_val = self.entry_phone.get().strip()

        if not email_val or not pw_val:
            messagebox.showerror("Missing fields", "Please enter both email and password.")
            return

        if not is_valid_email(email_val):
            messagebox.showerror("Invalid email", "Please enter a valid email address.")
            return

        # Normalize phone to dashed format before returning
        phone_digits = normalize_phone_digits(phone_val)
        phone_formatted = format_phone_dashed(phone_digits) if phone_digits else ""

        # Short network test in background
        def test_credentials():
            try:
                with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10) as smtp:
                    smtp.login(email_val, pw_val)
                # call success on main thread and pass phone_formatted
                self.top.after(0, lambda: self._finish_success(email_val, pw_val, phone_formatted))
            except smtplib.SMTPAuthenticationError:
                self.top.after(0, lambda: messagebox.showerror("Auth failed", "Authentication failed. Use an App Password for Gmail."))
            except Exception as e:
                self.top.after(0, lambda: messagebox.showerror("Connection error", f"Could not connect to SMTP server:\n{e}"))
            finally:
                self.top.after(0, lambda: self.status_label.configure(text=""))

        self.status_label.configure(text="Verifying credentials…")
        threading.Thread(target=test_credentials, daemon=True).start()

    def _finish_success(self, email_val, pw_val, phone_val):
        # Return a tuple (email, password, phone)
        self.result = (email_val, pw_val, phone_val)
        self.top.destroy()

    def _on_cancel(self):
        self.result = None
        self.top.destroy()

    def show(self):
        self.top.wait_window()
        return self.result

# -----------------------
# Main App
# -----------------------
class AutoMailCTKApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auto Mail Center — V4.1.0")
        self.geometry("1100x760")
        self.minsize(980, 680)

        # runtime state
        self.email_address = None
        self.email_password = None
        self.phone_number = None
        self.check_interval_seconds = DEFAULT_CHECK_INTERVAL
        self.replied_to = set()
        self._worker_thread = None
        self._stop_event = threading.Event()
        self._lock = threading.Lock()

        self._build_ui()
        self.after(120, self._do_login)

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar
        sidebar = ctk.CTkFrame(self, width=300, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nswe")
        sidebar.grid_rowconfigure(9, weight=1)

        logo = ctk.CTkLabel(sidebar, text="Auto Mail Center", font=ctk.CTkFont(size=18, weight="bold"))
        logo.pack(pady=(18, 6))

        self.lbl_logged_in = ctk.CTkLabel(sidebar, text="Not logged in", font=ctk.CTkFont(size=11))
        self.lbl_logged_in.pack(padx=12, pady=(6, 12))

        # Start/Stop
        self.btn_start = ctk.CTkButton(sidebar, text="▶ Start Auto-Responder", width=240, height=44, command=self._toggle_running)
        self.btn_start.pack(pady=(6, 12))

        # Interval control
        int_label = ctk.CTkLabel(sidebar, text="Check Interval (sec)", font=ctk.CTkFont(size=11))
        int_label.pack(pady=(10, 0))
        self.interval_slider = ctk.CTkSlider(sidebar, from_=5, to=600, width=240, command=self._on_interval_changed)
        self.interval_slider.set(self.check_interval_seconds)
        self.interval_slider.pack(pady=(6, 8))
        self.interval_value_label = ctk.CTkLabel(sidebar, text=str(self.check_interval_seconds))
        self.interval_value_label.pack()

        # Buttons
        ctk.CTkButton(sidebar, text="App Password Help", width=240, command=lambda: webbrowser.open("https://support.google.com/accounts/answer/185833")).pack(pady=(12, 4))
        ctk.CTkButton(sidebar, text="Logout", width=240, fg_color="#ef4444", hover_color="#f87171", command=self._logout).pack(pady=(6, 4))

        # Right content
        content = ctk.CTkFrame(self, corner_radius=8)
        content.grid(row=0, column=1, sticky="nswe", padx=12, pady=12)
        content.grid_rowconfigure(1, weight=1)
        content.grid_columnconfigure(0, weight=1)

        # Top bar
        top_bar = ctk.CTkFrame(content, fg_color="transparent")
        top_bar.grid(row=0, column=0, sticky="we", pady=(0, 12))
        top_bar.grid_columnconfigure(1, weight=1)

        self.status_label = ctk.CTkLabel(top_bar, text="Status: STOPPED", font=ctk.CTkFont(size=12, weight="bold"), text_color="#ef4444")
        self.status_label.grid(row=0, column=0, padx=6, sticky="w")

        self.last_replied_var = tk.StringVar(value="No replies yet")
        last_replied_lbl = ctk.CTkLabel(top_bar, textvariable=self.last_replied_var, font=ctk.CTkFont(size=11))
        last_replied_lbl.grid(row=0, column=1, sticky="e", padx=6)

        # Tabs
        tabs = ctk.CTkTabview(content, width=600)
        tabs.grid(row=1, column=0, sticky="nswe")
        tabs.add("Dashboard")
        tabs.add("Composer")
        tabs.add("Templates")
        tabs.add("System Log")
        tabs.set("Dashboard")

        # Dashboard
        dash = tabs.tab("Dashboard")
        ctk.CTkLabel(dash, text="Dashboard", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=(8, 4), padx=12)
        info_frame = ctk.CTkFrame(dash)
        info_frame.pack(padx=12, pady=6, fill="x")
        info_frame.grid_columnconfigure(0, weight=1)
        info_frame.grid_columnconfigure(1, weight=1)

        self.lbl_total_replies = ctk.CTkLabel(info_frame, text="Replies sent: 0", font=ctk.CTkFont(size=12))
        self.lbl_total_replies.grid(row=0, column=0, padx=12, pady=12, sticky="w")

        self.lbl_last_check = ctk.CTkLabel(info_frame, text="Last check: N/A", font=ctk.CTkFont(size=12))
        self.lbl_last_check.grid(row=0, column=1, padx=12, pady=12, sticky="e")

        # Composer
        composer = tabs.tab("Composer")
        ctk.CTkLabel(composer, text="Send Greeting", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=12, pady=(8, 6))

        form = ctk.CTkFrame(composer)
        form.pack(padx=12, pady=6, fill="x")

        ctk.CTkLabel(form, text="Recipient Name", font=ctk.CTkFont(size=11)).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="w")
        self.entry_recipient_name = ctk.CTkEntry(form, width=520)
        self.entry_recipient_name.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="w")

        ctk.CTkLabel(form, text="Recipient Email", font=ctk.CTkFont(size=11)).grid(row=2, column=0, padx=12, pady=(6, 4), sticky="w")
        self.entry_recipient_email = ctk.CTkEntry(form, width=520)
        self.entry_recipient_email.grid(row=3, column=0, padx=12, pady=(0, 8), sticky="w")

        ctk.CTkLabel(form, text="Template", font=ctk.CTkFont(size=11)).grid(row=4, column=0, padx=12, pady=(6, 4), sticky="w")
        self.template_optionmenu = ctk.CTkOptionMenu(form, values=list(GREETING_TEMPLATES.keys()), width=520)
        self.template_optionmenu.set(list(GREETING_TEMPLATES.keys())[0])
        self.template_optionmenu.grid(row=5, column=0, padx=12, pady=(0, 8), sticky="w")

        send_btn = ctk.CTkButton(form, text="📧 Send Greeting", width=220, command=self._on_send_greeting)
        send_btn.grid(row=6, column=0, padx=12, pady=(8, 16), sticky="w")

        # Templates
        templates = tabs.tab("Templates")
        ctk.CTkLabel(templates, text="Templates", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=12, pady=(8, 6))
        self.txt_template_preview = tk.Text(templates, height=16, wrap="word", bg="#07101a", fg="#dbeafe", insertbackground="#dbeafe")
        self.txt_template_preview.pack(padx=12, pady=8, fill="both", expand=True)
        self._update_template_preview()

        # System Log
        syslog = tabs.tab("System Log")
        ctk.CTkLabel(syslog, text="System Log", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=12, pady=(8, 6))
        self.txt_log = tk.Text(syslog, height=18, wrap="none", bg="#05060a", fg="#d6e3ff", insertbackground="#d6e3ff")
        self.txt_log.pack(padx=12, pady=8, fill="both", expand=True)

        # Bind template selection
        self.template_optionmenu.configure(command=lambda val: self._update_template_preview(val))

    # -------------------
    # Login
    # -------------------
    def _do_login(self):
        dlg = CTKLoginDialog(self)
        creds = dlg.show()
        if not creds:
            # user cancelled -> quit app cleanly
            self.destroy()
            return

        # creds is (email, password, phone)
        self.email_address, self.email_password, self.phone_number = creds
        phone_display = f"\nPhone: {self.phone_number}" if self.phone_number else ""
        self.lbl_logged_in.configure(text=f"Logged in as:\n{self.email_address}{phone_display}")
        self._push_log("✓ Login successful")
        self._push_log("Ready. Start the auto-responder when ready.")

    # -------------------
    # Interval change
    # -------------------
    def _on_interval_changed(self, val):
        try:
            self.check_interval_seconds = int(float(val))
        except Exception:
            self.check_interval_seconds = DEFAULT_CHECK_INTERVAL
        self.interval_value_label.configure(text=str(self.check_interval_seconds))

    # -------------------
    # Start / Stop
    # -------------------
    def _toggle_running(self):
        if not self._worker_thread or not self._worker_thread.is_alive():
            # start worker
            if not self.email_address or not self.email_password:
                messagebox.showerror("Not logged in", "Please login first.")
                return

            # clear stop event and start thread
            self._stop_event.clear()
            self._worker_thread = threading.Thread(target=self._worker_loop, daemon=True)
            self._worker_thread.start()
            self._set_status_running(True)
            self._push_log("▶ Auto-responder started")
        else:
            # request stop
            self._stop_event.set()
            self._set_status_running(False)
            self._push_log("⏸ Stop requested — waiting for worker to exit")

    def _set_status_running(self, running: bool):
        if running:
            self.status_label.configure(text="Status: RUNNING", text_color="#34d399")
            self.btn_start.configure(text="⏸ Stop Auto-Responder", fg_color="#ef4444", hover_color="#f47272")
        else:
            self.status_label.configure(text="Status: STOPPED", text_color="#ef4444")
            try:
                self.btn_start.configure(text="▶ Start Auto-Responder", fg_color="")
            except Exception:
                self.btn_start.configure(text="▶ Start Auto-Responder")

    # -------------------
    # Worker loop (thread)
    # -------------------
    def _worker_loop(self):
        """
        Worker that periodically checks inbox. Uses _stop_event to exit cleanly.
        """
        while not self._stop_event.is_set():
            # schedule logging onto main thread (safe)
            self._push_log("🔍 Checking inbox…")
            try:
                self._check_inbox_once()
            except Exception as e:
                self._push_log(f"❌ Worker error: {e}")
            # wait with early exit ability
            total = max(1, int(getattr(self, "check_interval_seconds", DEFAULT_CHECK_INTERVAL)))
            for _ in range(total):
                if self._stop_event.is_set():
                    break
                time.sleep(1)
        # worker exiting
        self._push_log("🛑 Worker exited cleanly")
        # ensure UI shows stopped (schedule on main thread)
        self.after(0, lambda: self._set_status_running(False))
        # clear worker reference
        self._worker_thread = None

    # -------------------
    # Check inbox once
    # -------------------
    def _check_inbox_once(self):
        """
        One-shot inbox check. Runs inside worker thread.
        """
        with self._lock:
            try:
                mail = imaplib.IMAP4_SSL("imap.gmail.com")
                # if credentials were cleared mid-check, bail out early
                if not self.email_address or not self.email_password:
                    try:
                        mail.logout()
                    except Exception:
                        pass
                    return
                mail.login(self.email_address, self.email_password)
                mail.select("inbox")
                status, data = mail.search(None, "UNSEEN")
                if status != 'OK':
                    self.after(0, lambda: self._push_log(f"⚠ IMAP search failed: {status}"))
                    mail.logout()
                    return

                id_list = data[0].split()
                if not id_list:
                    self.after(0, lambda: self._push_log("📭 No new messages"))
                    mail.logout()
                    self.after(0, lambda: self.lbl_last_check.configure(text=f"Last check: {datetime.now().strftime('%H:%M:%S')}"))
                    return

                for eid in id_list:
                    if self._stop_event.is_set():
                        break
                    try:
                        status, msg_data = mail.fetch(eid, "(RFC822)")
                        if status != 'OK':
                            self.after(0, lambda eid=eid: self._push_log(f"⚠ Failed to fetch id {eid}"))
                            continue

                        raw_msg = email.message_from_bytes(msg_data[0][1])
                        sender_hdr = raw_msg.get("From", "")
                        sender_email = email.utils.parseaddr(sender_hdr)[1]

                        if sender_email:
                            if sender_email in self.replied_to:
                                self.after(0, lambda se=sender_email: self._push_log(f"⏭ Already replied to {se}"))
                            else:
                                # send reply synchronously in worker (so we don't spawn too many threads)
                                self._send_auto_reply_internal(sender_email)
                        else:
                            self.after(0, lambda: self._push_log(f"⚠ Could not parse sender from: {sender_hdr}"))

                        mail.store(eid, '+FLAGS', '\\Seen')
                    except Exception as e:
                        self.after(0, lambda e=e: self._push_log(f"⚠ Error processing message: {e}"))
                try:
                    mail.close()
                except Exception:
                    pass
                mail.logout()
                self.after(0, lambda: self.lbl_last_check.configure(text=f"Last check: {datetime.now().strftime('%H:%M:%S')}"))
            except Exception as e:
                # Log on main thread
                self.after(0, lambda: self._push_log(f"❌ Inbox error: {e}"))

    # -------------------
    # Send auto-reply internal (called from worker thread)
    # -------------------
    def _send_auto_reply_internal(self, to_address):
        if to_address in self.replied_to:
            self.after(0, lambda: self._push_log(f"⏭ Already replied to {to_address}"))
            return

        msg = MIMEMultipart("alternative")
        msg['Subject'] = "Auto Reply"
        msg['From'] = self.email_address
        msg['To'] = to_address

        plain_text = ("Hi,\n\nThanks for your message – this is an automated reply "
                      "to confirm we've received it.\n\nBest,\nA.Apiwish")
        msg.attach(MIMEText(plain_text, "plain"))
        # Use phone number in HTML reply; fallback to default if not set
        phone_for_reply = self.phone_number or "000-000-0000"
        html = AUTO_REPLY_HTML_TEMPLATE.format(phone=phone_for_reply)
        msg.attach(MIMEText(html, "html"))

        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=20) as smtp:
                smtp.login(self.email_address, self.email_password)
                smtp.send_message(msg)
            self.replied_to.add(to_address)
            self.after(0, lambda: self._push_log(f"✓ Auto-reply sent to {to_address}"))
            self.after(0, lambda: self.lbl_total_replies.configure(text=f"Replies sent: {len(self.replied_to)}"))
            self.after(0, lambda: self.last_replied_var.set(f"Last replied: {to_address}"))
        except Exception as e:
            self.after(0, lambda: self._push_log(f"❌ ERROR auto-reply to {to_address}: {e}"))

    # -------------------
    # Manual greeting sender (UI thread triggers background worker)
    # -------------------
    def _on_send_greeting(self):
        to_addr = self.entry_recipient_email.get().strip()
        if not to_addr:
            messagebox.showerror("Missing email", "Please enter recipient email.")
            return
        if not is_valid_email(to_addr):
            messagebox.showerror("Invalid email", "Please enter a valid email address.")
            return

        recipient_name = self.entry_recipient_name.get().strip() or "there"
        template_name = self.template_optionmenu.get()
        template = GREETING_TEMPLATES.get(template_name, GREETING_TEMPLATES[list(GREETING_TEMPLATES.keys())[0]])
        html_content = template.replace("\n", "<br>").format(name=recipient_name)
        plain_content = template.format(name=recipient_name)

        def worker():
            try:
                msg = MIMEMultipart("alternative")
                msg['Subject'] = "Greeting from A.Apiwish"
                msg['From'] = self.email_address
                msg['To'] = to_addr
                msg.attach(MIMEText(plain_content, "plain"))
                msg.attach(MIMEText(html_content, "html"))

                with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=20) as smtp:
                    smtp.login(self.email_address, self.email_password)
                    smtp.send_message(msg)

                self.after(0, lambda: self._push_log(f"✓ Greeting sent to {to_addr} ({recipient_name})"))
                self.after(0, lambda: messagebox.showinfo("Success", f"Greeting email sent to {to_addr}!"))
            except Exception as e:
                self.after(0, lambda: self._push_log(f"❌ ERROR sending greeting: {e}"))
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to send email:\n{e}"))

        threading.Thread(target=worker, daemon=True).start()

    # -------------------
    # Template preview
    # -------------------
    def _update_template_preview(self, val=None):
        selected = val or self.template_optionmenu.get()
        text = GREETING_TEMPLATES.get(selected, "")
        self.txt_template_preview.delete("1.0", tk.END)
        self.txt_template_preview.insert(tk.END, text)

    # -------------------
    # Logging helper (UI thread safe)
    # -------------------
    def _push_log(self, text: str):
        """
        Thread-aware logging helper. If called from a background thread,
        schedule the log insertion on the main thread.
        """
        if threading.current_thread() is not threading.main_thread():
            # schedule to main thread
            try:
                self.after(0, lambda: self._push_log(text))
            except Exception:
                # fallback if widget destroyed
                print(f"(fallback) {text}", file=sys.stderr)
            return

        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        line = f"{ts} | {text}\n"
        try:
            self.txt_log.insert(tk.END, line)
            self.txt_log.see(tk.END)
        except Exception:
            # if UI is destroyed, just print to stdout for debugging
            print(line, file=sys.stderr)

    # -------------------
    # Logout (safe)
    # -------------------
    def _logout(self):
        # Ask user
        if messagebox.askyesno("Logout", "Are you sure you want to logout?"):
            # 1) Stop worker thread if running
            self._stop_event.set()  # request worker to stop
            self._push_log("🔒 Logout requested — stopping background worker")

            # 2) Clear credentials and in-memory caches (safely)
            self.email_address = None
            self.email_password = None
            self.phone_number = None
            self.replied_to.clear()

            # 3) Reset UI state
            self.lbl_logged_in.configure(text="Not logged in")
            self.lbl_total_replies.configure(text="Replies sent: 0")
            self.last_replied_var.set("No replies yet")
            self.lbl_last_check.configure(text="Last check: N/A")
            self._set_status_running(False)
            self._push_log("✓ Logged out (UI reset).")

            # 4) Wait until worker actually exits (non-blocking) then prompt login again
            def _wait_for_worker_exit():
                # if thread still alive, check again shortly
                if self._worker_thread and self._worker_thread.is_alive():
                    self.after(200, _wait_for_worker_exit)
                else:
                    # ensure reference cleared
                    self._worker_thread = None
                    # now prompt login (on main thread)
                    self.after(0, self._do_login)

            self.after(200, _wait_for_worker_exit)

    # -------------------
    # Clean exit on close
    # -------------------
    def destroy(self):
        # Request worker exit and give a brief moment
        try:
            self._stop_event.set()
            if self._worker_thread and self._worker_thread.is_alive():
                # allow short wait so worker can clean up
                self._worker_thread.join(timeout=1.0)
        except Exception:
            pass
        super().destroy()

# -----------------------
# Run
# -----------------------
if __name__ == "__main__":
    app = AutoMailCTKApp()
    app.mainloop()
