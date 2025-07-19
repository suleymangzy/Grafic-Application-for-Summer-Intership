import sys  # Sistem seviyesinde işlemler için (örn. argümanlar, çıkış)
import logging  # Hata ve olay günlüğü kaydı için

from PyQt5.QtWidgets import (  # PyQt5 arayüz öğeleri
    QApplication,  # Uygulama nesnesi (olmazsa olmaz)
    QMessageBox    # Hata mesaj kutusu (pop-up uyarı göstermek için)
)

from ui.mainWindow import MainWindow  # Uygulamanın ana penceresi (arayüz sınıfı)

# Ana çalıştırma bloğu: Bu dosya doğrudan çalıştırıldığında devreye girer
if __name__ == "__main__":
    app = QApplication(sys.argv)        # QApplication nesnesi oluşturulur, argv ile komut satırı argümanları alınır
    app.setStyle("Fusion")              # Fusion stili kullanılır (daha modern ve düz bir görünüm sağlar)

    try:
        win = MainWindow()              # MainWindow sınıfından ana pencere nesnesi oluşturulur
        win.show()                      # Ana pencere gösterilir
        sys.exit(app.exec_())           # Uygulama ana döngüsüne girilir ve uygulama çalışır
    except Exception as e:              # Eğer başlatma sırasında bir hata oluşursa
        logging.exception("Uygulama başlatılırken kritik bir hata oluştu.")  # Hata loglanır (stack trace ile)

        msg = QMessageBox()             # PyQt5 mesaj kutusu oluşturulur
        msg.setIcon(QMessageBox.Critical)  # Kritik hata ikonu (kırmızı çarpı) kullanılır
        msg.setText("Uygulama başlatılırken kritik bir hata oluştu.")  # Ana hata mesajı kullanıcıya gösterilir
        msg.setInformativeText(str(e))  # Hatanın detay metni (teknik açıklama)
        msg.setWindowTitle("Kritik Hata")  # Mesaj kutusu başlığı
        msg.exec_()                     # Mesaj kutusu çalıştırılır ve kullanıcıya gösterilir
        sys.exit(1)                     # Hata durumunda uygulama 1 kodu ile sonlandırılır
