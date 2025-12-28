"""
+================================================+
|                                                |
|                                                |
|                                                |
|     __           .__   __                      |
|   _/  |_ ___  ___|__|_/  |_ ___  ___  ______   |
|   \   __\\  \/  /|  |\   __\\  \/  / /  ___/   |
|    |  |   >    < |  | |  |   >    <  \___ \    |
|    |__|  /__/\_ \|__| |__|  /__/\_ \/____  >   |
|                \/                 \/     \/    |
|                                                |
|                                                |
|                                                |
+================================================+
"""
import os
import re
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import PhotoImage, Text, Tk, filedialog, messagebox
from tkinter import ttk


EMAIL_REGEX = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"


def is_valid_email(address: str) -> bool:
    cleaned = address.strip()
    return bool(re.fullmatch(EMAIL_REGEX, cleaned))


def validate_recipients(raw: str):
    addresses = [addr.strip() for addr in raw.split(",") if addr.strip()]
    invalid = [addr for addr in addresses if not is_valid_email(addr)]
    return addresses, invalid


def send_email(sender_email, sender_password, receiver_emails, subject, message_body,
               attachment_paths, email_provider, email_provider_port_number):
    smtp_server = f"smtp.{email_provider}"
    smtp_port = email_provider_port_number
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)

    for receiver_email in receiver_emails:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg["Subject"] = subject
        msg.attach(MIMEText(message_body, "plain"))

        for attachment_path in attachment_paths:
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename={os.path.basename(attachment_path)}",
            )
            msg.attach(part)

        server.sendmail(sender_email, receiver_email, msg.as_string())

    server.quit()


def build_ui():
    window = Tk()
    window.title("Email Automation Application")
    window.configure(bg="#0f172a")
    try:
        icon = PhotoImage(file="mail-inbox-app.png")
        window.iconphoto(True, icon)
    except Exception:
        pass

    # Consistent styling for a clean, readable UI
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TLabel", foreground="#e2e8f0", background="#0f172a", font=("Segoe UI", 11))
    style.configure("TEntry", fieldbackground="#1e293b", foreground="#e2e8f0")
    style.configure(
        "Input.TEntry",
        padding=(8, 6),
        relief="flat",
        fieldbackground="#111827",
        foreground="#e2e8f0",
        insertcolor="#e2e8f0",
    )
    style.map(
        "Input.TEntry",
        fieldbackground=[("focus", "#0b1221"), ("!focus", "#111827")],
    )
    style.configure("TButton", font=("Segoe UI", 11, "bold"), padding=10)
    style.map("TButton", background=[("active", "#0ea5e9"), ("!active", "#0284c7")], foreground=[("active", "#0f172a"), ("!active", "#e2e8f0")])
    style.configure("Card.TFrame", background="#111827", bordercolor="#1f2937", relief="ridge", borderwidth=1)

    header = ttk.Frame(window, style="Card.TFrame", padding=20)
    header.pack(fill="x", padx=20, pady=(20, 10))
    ttk.Label(header, text="Email Automation", font=("Segoe UI Semibold", 20)).pack(anchor="w")
    ttk.Label(header, text="Sende schnell formatierte Mails mit Anhängen.", font=("Segoe UI", 12), foreground="#94a3b8").pack(anchor="w", pady=(4, 0))

    form = ttk.Frame(window, style="Card.TFrame", padding=20)
    form.pack(fill="both", expand=True, padx=20, pady=10)

    sender_email_var = ttk.Entry(form, width=40)
    sender_email_var.insert(0, "you@example.com")
    sender_password_var = ttk.Entry(form, width=40, show="*")
    receiver_emails_var = ttk.Entry(form, width=60)
    subject_var = ttk.Entry(form, width=60)
    body_text = Text(form, height=12, width=90, bg="#0b1221", fg="#e2e8f0", insertbackground="#e2e8f0", highlightbackground="#1f2937", relief="flat", wrap="word")
    attachment_list = []

    ttk.Label(form, text="Absender").grid(row=0, column=0, sticky="w", pady=(0, 6))
    sender_email_var.grid(row=1, column=0, sticky="w", pady=(0, 14))
    ttk.Label(form, text="Passwort").grid(row=0, column=1, sticky="w", padx=(20, 0), pady=(0, 6))
    sender_password_var.grid(row=1, column=1, sticky="w", padx=(20, 0), pady=(0, 14))

    ttk.Label(form, text="Empfänger (Komma-getrennt)").grid(row=2, column=0, sticky="w", pady=(0, 6))
    receiver_emails_var.grid(row=3, column=0, columnspan=2, sticky="we", pady=(0, 14))

    ttk.Label(form, text="Betreff").grid(row=4, column=0, sticky="w", pady=(0, 6))
    subject_var.grid(row=5, column=0, columnspan=2, sticky="we", pady=(0, 14))

    ttk.Label(form, text="Text").grid(row=6, column=0, sticky="w", pady=(0, 6))
    body_text.grid(row=7, column=0, columnspan=2, sticky="we", pady=(0, 14))

    ttk.Label(form, text="Anhänge").grid(row=8, column=0, sticky="w", pady=(0, 6))
    attachments_box = Text(form, height=5, width=90, bg="#0b1221", fg="#e2e8f0", state="disabled", relief="flat")
    attachments_box.grid(row=9, column=0, columnspan=2, sticky="we", pady=(0, 10))

    def refresh_attachments_box():
        attachments_box.configure(state="normal")
        attachments_box.delete("1.0", "end")
        if attachment_list:
            attachments_box.insert("1.0", "\n".join(attachment_list))
        attachments_box.configure(state="disabled")

    def add_attachments():
        files = filedialog.askopenfilenames(title="Anhänge auswählen")
        for file_path in files:
            if file_path and file_path not in attachment_list:
                attachment_list.append(file_path)
        refresh_attachments_box()

    def clear_attachments():
        attachment_list.clear()
        refresh_attachments_box()

    buttons_row = ttk.Frame(form, style="Card.TFrame")
    buttons_row.grid(row=10, column=0, columnspan=2, sticky="w", pady=(0, 14))
    ttk.Button(buttons_row, text="Anhänge hinzufügen", command=add_attachments).pack(side="left", padx=(0, 10))
    ttk.Button(buttons_row, text="Anhänge leeren", command=clear_attachments).pack(side="left")

    provider_row = ttk.Frame(form, style="Card.TFrame", padding=(0, 6, 0, 0))
    provider_row.grid(row=11, column=0, columnspan=2, sticky="w")
    ttk.Label(provider_row, text="Provider (z.B. gmail.com)").pack(side="left", padx=(0, 10))
    provider_var = ttk.Entry(provider_row, width=26, style="Input.TEntry")
    provider_var.insert(0, "gmail.com")
    provider_var.pack(side="left", padx=(0, 20))
    ttk.Label(provider_row, text="Port").pack(side="left", padx=(0, 8))
    port_var = ttk.Entry(provider_row, width=8, style="Input.TEntry")
    port_var.insert(0, "587")
    port_var.pack(side="left")

    status_var = ttk.Label(window, text="Bereit", anchor="w")
    status_var.pack(fill="x", padx=20, pady=(0, 10))

    def update_status(message, is_error=False):
        status_var.configure(text=message, foreground="#f87171" if is_error else "#34d399")
        window.update_idletasks()

    def submit_email_info():
        sender_email = sender_email_var.get().strip()
        email_password = sender_password_var.get()
        receivers, invalid_receivers = validate_recipients(receiver_emails_var.get())
        subject = subject_var.get().strip()
        email_body = body_text.get("1.0", "end").strip()
        email_provider = provider_var.get().strip()

        try:
            port_int = int(port_var.get().strip())
        except ValueError:
            update_status("Port muss eine Zahl sein.", True)
            return

        if not is_valid_email(sender_email):
            update_status("Absender-Adresse ist ungültig.", True)
            return

        if invalid_receivers:
            update_status(f"Ungültige Empfänger: {', '.join(invalid_receivers)}", True)
            return

        if not receivers:
            update_status("Mindestens ein Empfänger wird benötigt.", True)
            return

        if not subject:
            update_status("Betreff darf nicht leer sein.", True)
            return

        missing = [p for p in attachment_list if not os.path.isfile(p)]
        if missing:
            update_status(f"Datei nicht gefunden: {missing[0]}", True)
            return

        update_status("Sende...", False)
        try:
            send_email(sender_email, email_password, receivers, subject, email_body,
                       attachment_list, email_provider, port_int)
            update_status("E-Mail erfolgreich gesendet.", False)
            messagebox.showinfo("Erfolg", "E-Mail wurde versendet.")
        except Exception as error:  # smtp errors
            update_status(f"Fehler: {error}", True)

    action_bar = ttk.Frame(window, style="Card.TFrame", padding=16)
    action_bar.pack(fill="x", padx=20, pady=10)
    ttk.Button(action_bar, text="E-Mail senden", command=submit_email_info).pack(side="right")

    form.columnconfigure(0, weight=1)
    form.columnconfigure(1, weight=1)

    # Derive initial size from required space, enforce a minimum, allow resize.
    window.update_idletasks()
    padding = 60
    req_width = window.winfo_reqwidth() + padding
    req_height = window.winfo_reqheight() + padding
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    width = min(req_width, screen_w)
    height = min(req_height, screen_h)
    x = (screen_w - width) // 2
    y = (screen_h - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")
    window.minsize(width, height)
    window.resizable(True, True)

    window.mainloop()


if __name__ == "__main__":
    build_ui()