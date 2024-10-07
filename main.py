"""Imports"""
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import logging
import re
import smtplib
from textwrap import fill
import threading
import tkinter as tk
from tkinter import PhotoImage, font, filedialog, messagebox
from tkinter.ttk import Button, Entry, Frame, Label, Progressbar, Style
import sys

from numpy import pad
import pandas as pd


class SendMailsApp:
    """Aplicação para envio de e-mails"""

    def __init__(self, main_window):
        """Settings Interface Application"""
        self.main_window = main_window
        self.main_window.title("SendMails")
        self.main_window.resizable(False, False)

        # Window settings
        window_width = 600
        window_height = 750
        screen_width = self.main_window.winfo_screenwidth()
        screen_height = self.main_window.winfo_screenheight()
        center_x = int((screen_width / 2) - (window_width / 2))
        center_y = int((screen_height / 2) - (window_height / 2))
        self.main_window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

        self.setup_logging()
        self.set_app_icon()
        self.create_styles()
        self.create_widgets()
        
    def setup_logging(self):
        """Configure LOG file"""
        log_filename = "LOG_SENDMAILS.log"
        logging.basicConfig(
            filename=log_filename,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        self.logger = logging.getLogger()
        
    def set_app_icon(self):
        """Define o icone da aplicação dependendo do sistema operacional"""
        try:
            if sys.platform.startswith('win'):
                self.main_window.iconbitmap('./icon.ico')
            else:
                icon = PhotoImage(file='./icon.png')
                self.main_window.iconphoto(True, icon)
                self.app_icon = icon
        except Exception as e:
            self.logger.error(f"Erro ao definir o ícone da aplicação: {e}")

    def create_styles(self):
        """Create Styles"""
        self.font_default = font.Font(family="Quicksand Medium", size=10)
        self.main_window.option_add("*Font", self.font_default)

        self.style = Style()
        self.style.theme_use('default')
        self.style.configure('TFrame', background="#edf2f4", foreground="#212529")
        self.style.configure('TLabel', background="#edf2f4", foreground="#212529")
        self.style.configure('TEntry', padding=10, font=("Quicksand Medium", 12))
        self.style.configure('TButton', padding=10, font=("Quicksand Bold", 10))
        self.style.configure('send.TButton', background="#5cb85c", foreground="#FFFFFF")
        self.style.map("send.TButton", background=[('active', '#4cae4c')])
        self.style.configure('clear.TButton', background="#f0ad4e", foreground="#FFFFFF")
        self.style.map("clear.TButton", background=[('active', '#ec971f')])
        self.style.configure('file.TButton', background="#0275d8", foreground="#FFFFFF")
        self.style.map("file.TButton", background=[('active', '#025aa5')])
        self.style.configure('exit.TButton', background="#d9534f", foreground="#FFFFFF")
        self.style.map("exit.TButton", background=[('active', '#c9302c')])
        self.style.configure("custom.Horizontal.TProgressbar", background="#5cb85c", troughcolor='lightgray')

    def create_widgets(self):
        """Creates the main interface"""
        body_frame = Frame(self.main_window)
        body_frame.pack(padx=20, pady=10, fill='both', expand=True)

        title_label = Label(body_frame, text="SendMails", font=("Quicksand Bold", 24))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky='n')

        subject_label = Label(body_frame, text="Assunto")
        subject_label.grid(row=1, column=0, sticky='w', padx=10, pady=(10, 0))
        self.subject_entry = Entry(body_frame, style="TEntry")
        self.subject_entry.grid(row=2, column=0, columnspan=2, sticky='ew', padx=10, pady=10)
        self.subject_entry.focus()

        message_label = Label(body_frame, text="Mensagem")
        message_label.grid(row=3, column=0, sticky='w', padx=10, pady=(10, 0))
        self.message_text = tk.Text(body_frame, height=10, width=50, padx=10, pady=10)
        self.message_text.grid(row=4, column=0, columnspan=2, sticky='ew', padx=10, pady=10)

        # Lista para armazenar referências a todos os botões
        self.buttons = []
        
        file_image_frame = Frame(body_frame)
        file_image_frame.grid(row=5, column=0, columnspan=2, sticky='ew', padx=10, pady=10)
        self.file_image_entry = Entry(file_image_frame, style="TEntry")
        self.file_image_entry.pack(side='left', fill='x', expand=True)
        self.btn_image = Button(file_image_frame, text="IMAGEM", style="file.TButton", command=self.select_image)
        self.btn_image.pack(side='right', padx=(10, 0))
        self.buttons.append(self.btn_image)

        excel_file_frame = Frame(body_frame)
        excel_file_frame.grid(row=6, column=0, columnspan=2, sticky='ew', padx=10, pady=10)
        self.file_excel_entry = Entry(excel_file_frame, style="TEntry")
        self.file_excel_entry.pack(side='left', fill='x', expand=True)
        self.btn_recipient = Button(excel_file_frame, text="E-MAILS", style="file.TButton", command=self.select_excel)
        self.btn_recipient.pack(side='right', padx=(10, 0))
        self.buttons.append(self.btn_recipient)

        spacer = Label(body_frame, text="")
        spacer.grid(row=7, column=0, columnspan=2, pady=5)

        buttons_frame = Frame(body_frame)
        buttons_frame.grid(row=8, column=0, columnspan=2, sticky='ew')
        buttons_frame.columnconfigure(0, weight=1)
        buttons_frame.columnconfigure(1, weight=1)
        buttons_frame.columnconfigure(2, weight=1)      

        self.btn_send = Button(buttons_frame, text="ENVIAR", style="send.TButton", command=self.start_sending_emails)
        self.btn_send.grid(row=0, column=0, padx=5, sticky='ew')
        self.buttons.append(self.btn_send)
        
        self.btn_clear = Button(buttons_frame, text="LIMPAR", style="clear.TButton", command=self.clear_fields)
        self.btn_clear.grid(row=0, column=1, padx=5, sticky='ew')
        self.buttons.append(self.btn_clear)

        self.btn_exit = Button(buttons_frame, text="SAIR", style="exit.TButton", command=self.main_window.destroy)
        self.btn_exit.grid(row=0, column=2, padx=5, sticky='ew')
        self.buttons.append(self.btn_exit)
        
        # Barra de progresso
        progress_label = Label(body_frame, text="Progresso de envio dos e-mails")
        progress_label.grid(row=9, column=0, sticky='w', padx=10, pady=(15, 5))
        progress_frame = Frame(body_frame)
        progress_frame.grid(row=10, column=0, columnspan=3, sticky='ew')
        self.progress_bar = Progressbar(progress_frame, style="custom.Horizontal.TProgressbar")
        self.progress_bar.pack(fill='x', expand=True, padx=10)
        
        # Author
        author_label = Label(body_frame, text="Desenvolvido por BRUNA GONÇALVES", font=("Quicksand Bold", 10))
        author_label.grid(row=11, columnspan=3, pady=(15, 0))        

        # Configurar expansão das colunas
        for i in range(2):
            body_frame.columnconfigure(i, weight=1)

    def select_image(self):
        """Select image file"""
        file_path = filedialog.askopenfilename(title="Abrir", filetypes=[("Arquivos de Imagem", "*.png *.jpg *.jpeg *.gif")])        
        if file_path:
            self.file_image_entry.delete(0, tk.END)
            self.file_image_entry.insert(0, file_path)
        else:
            self.file_image_entry.delete(0, tk.END)
            
    def select_excel(self):
        """Select excel file"""
        file_path = filedialog.askopenfilename(title="Abrir", filetypes=[("Arquivos Excel", "*.xlsx")])        
        if file_path:
            self.file_excel_entry.delete(0, tk.END)
            self.file_excel_entry.insert(0, file_path)
        else:
            self.file_excel_entry.delete(0, tk.END)
    
    def validate_fields(self):
        """Validate input fields before sending emails"""
        subject = self.subject_entry.get().strip()
        message = self.message_text.get("1.0", tk.END).strip()
        image = self.file_image_entry.get().strip()
        excel = self.file_excel_entry.get().strip()
        
        if not subject:
            messagebox.showerror("Erro de Validação",
                                 "O campo 'Assunto' é obrigatório.")
            return False
        if not excel:
            messagebox.showerror("Erro de Validação",
                                 "Obrigatório adicionar a lista de destinatários!")
            return False
        if not message and not image:
            messagebox.showerror(
                "Erro de Validação", "Pelo menos um dos campos 'Mensagem' ou 'Imagem' deve estar preenchido.")
            return False
        return True
        
    def send_emails(self):
        """Send mails based on the provided Excel list"""
        subject = self.subject_entry.get().strip()
        message = self.message_text.get("1.0", tk.END).strip()
        excel_path = self.file_excel_entry.get().strip()
        image_path = self.file_image_entry.get().strip()
        
        # Desabilitar o botão antes de iniciar o envio
        self.main_window.after(0, lambda: self.disable_buttons('disabled'))
        
        # Lista para armazenar e-mails com erros
        error_emails = []
        
        try:
            try:
                df = pd.read_excel(excel_path)
                if 'Emails' not in df.columns:
                    messagebox.showerror(
                        "Erro", "A coluna 'Emails' não foi encontrada no arquivo Excel.")
                    self.logger.error(
                        "A coluna 'Emails' não foi encontrada no arquivo Excel: %s", excel_path)
                    return
            except FileNotFoundError:
                messagebox.showerror(
                    "Erro", f"O arquivo Excel não foi encontrado: {excel_path}")
                self.logger.error(
                    "Arquivo Excel não encontrado: %s", excel_path)
                return
            except Exception as e:
                messagebox.showerror("Erro", str(e))
                self.logger.error(
                    "Erro ao ler o arquivo Excel: %s", str(e))
                return
            
            # Set up the SMTP server
            smtp_server = 'smtp.email.com'
            smtp_port = 587
            sender_email = 'email@email.com'
            sender_password = 'senha'

            try:
                server = smtplib.SMTP(smtp_server, smtp_port)
                server.starttls()
                server.login(sender_email, sender_password)
            except Exception as e:
                messagebox.showerror("Erro SMTP", f"Falha ao conectar ao servidor SMTP: {e}")
                self.logger.error("Falha ao conectar ao servidor SMTp: %s", str(e))
                return
                        
            total_emails = len(df)
            self.progress_bar['maximum'] = total_emails
            self.progress_bar['value'] = 0
            successful_sends = 0 
            for index, row in df.iterrows():
                recipient_email = row['Emails']
                if self.is_valid_email(recipient_email):
                    try:
                        self.logger.info("Enviando e-mail para: %s", recipient_email)
                        self.send_email(server, sender_email, recipient_email, subject, message, image_path)
                        successful_sends += 1
                    except Exception as e:
                        error_emails.append((recipient_email, f"Falha no envio: {str(e)}"))
                        self.logger.error("Falha ao enviar e-mail para %s: %s", recipient_email, str(e))
                else:
                    error_emails.append((recipient_email, "Formato de e-mail inválido"))
                    self.logger.warning("Endereço de e-mail inválido encontrado: %s", recipient_email)

                self.progress_bar['value'] += 1
                self.main_window.update_idletasks()
                
            server.quit()
            self.progress_bar['value'] = self.progress_bar['maximum']
            
            if error_emails:
                error_message = "Alguns e-mails não foram enviados:\n\n"
                for email, motivo in error_emails:
                    error_message += f"{email} - {motivo}\n"
                # Utilizar uma janela separada para exibir erros se a lista for muito longa
                if len(error_emails) > 10:
                    self.show_error_emails(error_emails)
                else:
                    messagebox.showwarning("E-mails com Erros", error_message)
                self.logger.info("E-mails com erros: %s", error_emails)
            else:
                messagebox.showinfo("Sucesso", "Todos os e-mails foram enviados com sucesso.")
                self.logger.info("Todos os e-mails foram enviados com sucesso. Total: %d", successful_sends)

            
        except Exception as e:
            self.main_window.after(0, lambda: messagebox.showerror("Erro", str(e)))
            self.logger.error("Erro ao enviar e-mails: %s", str(e))
        finally:
            self.main_window.after(0, lambda: self.disable_buttons('normal'))
        
    def send_email(self, server, sender_email, recipient_email, subject, message, image_path):
        """Function to send a single email"""
        msg = MIMEMultipart('related')
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email

        # Create the HTML message with inline image
        html_content = f"""
        <html>
        <body>
            <p>{message}</p>
            <img src="cid:image1">
        </body>
        </html>
        """
        msg.attach(MIMEText(html_content, 'html'))

        # Attach the image if provided
        if image_path:
            with open(image_path, 'rb') as img_file:
                # Use 'png' or 'gif' if needed
                img = MIMEImage(img_file.read(), 'jpeg')
                img.add_header('Content-ID', '<image1>')
                msg.attach(img)

        # Send the email
        server.sendmail(sender_email, recipient_email, msg.as_string())

    def start_sending_emails(self):
        """Start the email sending process in a separate thread"""
        if not self.validate_fields():
            return
        
        send_thread = threading.Thread(target=self.send_emails)
        send_thread.start()
        
    def disable_buttons(self, state):
        """Enable or disable the send button"""
        for button in self.buttons:
            button.config(state=state)
                
    def clear_fields(self):
        """Clear all input fields if they are filled"""
        # Verifica se algum campo está preenchido
        subject_filled = bool(self.subject_entry.get().strip())
        message_filled = bool(self.message_text.get("1.0", tk.END).strip())
        image_filled = bool(self.file_image_entry.get().strip())
        excel_filled = bool(self.file_excel_entry.get().strip())

        if subject_filled or message_filled or image_filled or excel_filled:
            confirm = messagebox.askyesno(
                "Confirmar Limpeza",
                "Tem certeza que deseja limpar todos os campos preenchidos?"
            )
            if confirm:
                self.subject_entry.delete(0, tk.END)
                self.message_text.delete("1.0", tk.END)
                self.file_image_entry.delete(0, tk.END)
                self.file_excel_entry.delete(0, tk.END)
                self.logger.info("Todos os campos foram limpos pelo usuário.")
                messagebox.showinfo("Campos Limpados", "Todos os campos foram limpos com sucesso.")
                self.progress_bar['value'] = 0
                self.main_window.update_idletasks()
        else:
            messagebox.showinfo("Nenhum Campo Preenchido", "Não há campos para limpar.")
    
    def is_valid_email(self, email):
        """Valida se o e-mail possui um formato válido."""
        regex = r'^\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        return re.match(regex, email)
    
    def show_error_emails(self, error_emails):
        """Exibir os e-mails com erros em uma nova janela"""
        error_window = tk.Toplevel(self.main_window)
        error_window.title("E-mails com Erros")
        error_window.geometry("500x400")
        
        # Adicionar uma barra de rolagem
        scrollbar = tk.Scrollbar(error_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Adicionar o widget Text
        text = tk.Text(error_window, wrap='word', yscrollcommand=scrollbar.set)
        text.pack(expand=True, fill='both')
        
        # Inserir os erros no widget Text
        for email, motivo in error_emails:
            text.insert(tk.END, f"{email} - {motivo}\n")
        
        # Configurar a barra de rolagem
        scrollbar.config(command=text.yview)
        
        # Tornar a janela modal
        error_window.transient(self.main_window)
        error_window.grab_set()
        self.main_window.wait_window(error_window)

    def update_progress_bar(self, value):
        """Update the progress bar value"""
        self.progress_bar['value'] = value
            
if __name__ == "__main__":
    window = tk.Tk()
    app = SendMailsApp(window)
    window.mainloop()
