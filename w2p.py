import os
import sys
import re
import win32com.client
from PyPDF2 import PdfMerger


# Yapılandırma ve Renk Kodları
class UI:
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    RESET = '\033[0m'

    @staticmethod
    def setup_terminal():
        os.system("")  # ANSI desteği
        if os.name == 'nt':
            os.system('mode con: cols=100 lines=30')

    @staticmethod
    def print_header():
        os.system('cls' if os.name == 'nt' else 'clear')
        columns = 100
        try:
            columns = os.get_terminal_size().columns
        except OSError:
            pass

        banner = [
            "██████╗ ██╗  ██╗ ██████╗ ",
            "██╔══██╗██║ ██╔╝██╔════╝ ",
            "██████╔╝█████╔╝ ██║  ███╗",
            "██╔══██╗██╔═██╗ ██║   ██║",
            "██████╔╝██║  ██╗╚██████╔╝",
            "╚══════╝╚═╝  ╚═╝ ╚═════╝ "
        ]
        print(f"\n{UI.CYAN}{UI.BOLD}")
        for line in banner:
            print(line.center(columns))

        tag = ">>> PDF ENGINE v2.0 | AUTOMATION MODULE <<<"
        print(f"{UI.GREEN}{tag.center(columns)}{UI.RESET}")
        print("-" * columns)


class PDFProcessor:
    def __init__(self):
        self.base_path = self._get_root()
        self.output_path = os.path.join(self.base_path, "output")
        self.extensions = (".docx", ".doc", ".rtf", ".odt", ".txt")

        if not os.path.exists(self.output_path):
            os.makedirs(self.output_path)

    def _get_root(self):
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _natural_sort(self, text):
        return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', text)]

    def convert_documents(self):
        print(f"{UI.YELLOW}[1/2] Belgeler PDF formatına dönüştürülüyor...{UI.RESET}")

        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
        except Exception as e:
            print(f"{UI.RED}Hata: MS Word başlatılamadı. {e}{UI.RESET}")
            return

        files = [f for f in os.listdir(self.base_path)
                 if f.lower().endswith(self.extensions) and not f.startswith("~$")]

        if not files:
            print("  - İşlenecek belge bulunamadı.")
            word.Quit()
            return

        for file in files:
            try:
                in_path = os.path.join(self.base_path, file)
                out_name = os.path.splitext(file)[0] + ".pdf"
                out_path = os.path.join(self.output_path, out_name)

                print(f"  > {file} dönüştürülüyor...")
                doc = word.Documents.Open(in_path)
                doc.SaveAs(out_path, FileFormat=17)
                doc.Close()
            except Exception as e:
                print(f"  ! Hata: {file} atlandı. ({e})")

        word.Quit()
        print(f"{UI.GREEN}[+] Dönüştürme tamamlandı.{UI.RESET}\n")

    def merge_pdfs(self):
        print(f"{UI.YELLOW}[2/2] PDF dosyaları birleştiriliyor...{UI.RESET}")

        # Önce output klasörüne bak
        current_dir = self.output_path
        pdfs = [f for f in os.listdir(self.output_path)
                if f.lower().endswith(".pdf") and f != "_FINAL_REPORT.pdf"]

        # Eğer output boşsa ana klasöre bak
        if not pdfs:
            current_dir = self.base_path
            pdfs = [f for f in os.listdir(self.base_path)
                    if f.lower().endswith(".pdf") and f != "_FINAL_REPORT.pdf"]

            if not pdfs:
                print(f"{UI.RED}[!] Birleştirilecek dosya bulunamadı.{UI.RESET}")
                return

        # Sıralama ve Birleştirme
        pdfs.sort(key=self._natural_sort)
        merger = PdfMerger()

        for pdf in pdfs:
            print(f"  + Listeye eklendi: {pdf}")
            # Dosyayı bulduğu klasörden (current_dir) çekmesini sağlıyoruz
            full_path = os.path.join(current_dir, pdf)
            merger.append(full_path)

        final_out = os.path.join(self.output_path, "_FINAL_REPORT.pdf")
        merger.write(final_out)
        merger.close()

        print(f"\n{UI.GREEN}{UI.BOLD}[BAŞARILI] Sonuç: {final_out}{UI.RESET}")


def main():
    UI.setup_terminal()
    UI.print_header()

    engine = PDFProcessor()
    engine.convert_documents()
    print("-" * 60)
    engine.merge_pdfs()

    print(f"\n{UI.CYAN}İşlem başarıyla sonlandırıldı.{UI.RESET}")
    input("Çıkış yapmak için Enter tuşuna basın...")


if __name__ == "__main__":
    main()