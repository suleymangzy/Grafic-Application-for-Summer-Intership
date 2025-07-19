from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QWidget,
    QPushButton,
    QLabel,
    QVBoxLayout,
    QHBoxLayout,
    QListWidget,
    QListWidgetItem,
    QComboBox,
    QMessageBox,
    QScrollArea,
    QCheckBox,
    QSpacerItem,
)


class DataSelectionPage(QWidget):
    """
    Günlük grafikler için veri seçimi arayüzünü temsil eden QWidget.
    Kullanıcıya Excel dosyasındaki sayfa seçimi, gruplanacak tarihler,
    gruplanan değişkenler (örneğin ürünler) ve metrikler için seçim yapma imkanı verir.
    """

    def __init__(self, main_window: "MainWindow") -> None:
        """Ana pencere referansını alır ve arayüzü başlatır."""
        super().__init__()
        self.main_window = main_window
        self.init_ui()

    def init_ui(self):
        """Arayüz bileşenlerini oluşturur ve düzenleri ayarlar."""
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignTop)

        # Başlık etiketi
        title_label = QLabel("<h2>Günlük Grafik Veri Seçimi</h2>")
        title_label.setObjectName("title_label")  # Stil için objectName verildi
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # Sayfa seçimi kısmı (Excel sayfaları listesi)
        sheet_selection_group = QHBoxLayout()
        self.sheet_selection_label = QLabel("İşlenecek Sayfa:")
        self.sheet_selection_label.setAlignment(Qt.AlignLeft)
        sheet_selection_group.addWidget(self.sheet_selection_label)

        self.cmb_sheet = QComboBox()
        self.cmb_sheet.setEnabled(False)  # Başlangıçta devre dışı
        self.cmb_sheet.currentIndexChanged.connect(self.on_sheet_selected)
        sheet_selection_group.addWidget(self.cmb_sheet)
        main_layout.addLayout(sheet_selection_group)

        # Gruplama değişkeni seçimi (örneğin tarihler)
        grouping_group = QHBoxLayout()
        grouping_group.addWidget(QLabel("<b>Gruplama Değişkeni (Tarihler):</b>"))
        self.cmb_grouping = QComboBox()
        self.cmb_grouping.currentIndexChanged.connect(self.populate_grouped)
        grouping_group.addWidget(self.cmb_grouping)
        main_layout.addLayout(grouping_group)

        # Gruplanan değişkenler (ürünler) listesi, çoklu seçim destekli
        grouped_group = QHBoxLayout()
        grouped_group.addWidget(QLabel("<b>Gruplanan Değişkenler (Ürünler):</b>"))
        self.lst_grouped = QListWidget()
        self.lst_grouped.setSelectionMode(QListWidget.MultiSelection)
        self.lst_grouped.itemSelectionChanged.connect(self.update_next_button_state)
        grouped_group.addWidget(self.lst_grouped)
        main_layout.addLayout(grouped_group)

        # Metrikler için onay kutuları (kaydırılabilir alan içinde)
        metrics_group = QVBoxLayout()
        metrics_group.addWidget(QLabel("<b>Metrikler :</b>"))
        self.metrics_scroll_area = QScrollArea()
        self.metrics_scroll_area.setWidgetResizable(True)
        self.metrics_content_widget = QWidget()
        self.metrics_layout = QVBoxLayout(self.metrics_content_widget)
        self.metrics_scroll_area.setWidget(self.metrics_content_widget)
        metrics_group.addWidget(self.metrics_scroll_area)
        main_layout.addLayout(metrics_group)

        # Navigasyon butonları (Geri, İleri)
        nav_layout = QHBoxLayout()
        self.btn_back = QPushButton("← Geri")
        self.btn_back.clicked.connect(lambda: self.main_window.goto_page(0))
        nav_layout.addWidget(self.btn_back)

        self.btn_next = QPushButton("İleri →")
        self.btn_next.setEnabled(False)  # Başlangıçta pasif
        self.btn_next.clicked.connect(self.go_next)
        nav_layout.addStretch(1)  # İleri butonunu sağa yasla
        nav_layout.addWidget(self.btn_next)
        main_layout.addLayout(nav_layout)

    def _populate_data_selection_fields(self):
        """
        DataFrame verilerini kullanarak gruplanacak tarihler,
        gruplanan değişkenler ve metrikler için seçim alanlarını doldurur.
        """
        df = self.main_window.df
        if df.empty:
            QMessageBox.critical(self, "Hata", "Veri yüklenemedi. Lütfen dosyayı kontrol edin.")
            self.main_window.goto_page(0)
            return

        # Gruplama comboBox sinyallerini geçici engelle (gereksiz tetiklemeyi önlemek için)
        self.cmb_grouping.blockSignals(True)
        self.cmb_grouping.clear()

        grouping_col_name = self.main_window.grouping_col_name
        if grouping_col_name and grouping_col_name in df.columns:
            # Gruplama sütunundaki eşsiz ve boş olmayan değerleri sırala
            grouping_vals = sorted(df[grouping_col_name].dropna().astype(str).unique())
            grouping_vals = [s for s in grouping_vals if s.strip()]  # boş stringleri çıkar
            self.cmb_grouping.addItems(grouping_vals)
            if not grouping_vals:
                QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) boş veya geçerli değer içermiyor.")
        else:
            # Gruplama sütunu yoksa veya geçersizse
            QMessageBox.warning(self, "Uyarı", "Gruplama sütunu (A) bulunamadı veya boş.")
            self.cmb_grouping.clear()
            self.lst_grouped.clear()
            self.clear_metrics_checkboxes()
            self.update_next_button_state()
            self.cmb_grouping.blockSignals(False)
            return

        # Sinyalleri tekrar etkinleştir
        self.cmb_grouping.blockSignals(False)

        # Metrik onay kutularını ve gruplanan öğeleri doldur
        self.populate_metrics_checkboxes()
        self.populate_grouped()

    def populate_grouped(self) -> None:
        """
        Seçilen gruplanma değerine göre (örn. seçilen tarih)
        gruplanan değişkenler listesini (örneğin ürünler) günceller ve seçili yapar.
        """
        self.lst_grouped.clear()
        selected_grouping_val = self.cmb_grouping.currentText()
        df = self.main_window.df

        if selected_grouping_val and self.main_window.grouping_col_name and self.main_window.grouped_col_name:
            # Seçilen gruplanma değerine göre dataframe'i filtrele
            filtered_df = df[df[self.main_window.grouping_col_name].astype(str) == selected_grouping_val]
            # Gruplanan sütundaki eşsiz ve boş olmayan değerleri sırala
            grouped_vals = sorted(filtered_df[self.main_window.grouped_col_name].dropna().astype(str).unique())
            grouped_vals = [s for s in grouped_vals if s.strip()]

            # Her değeri list widget'a ekle ve varsayılan seçili yap
            for gv in grouped_vals:
                item = QListWidgetItem(gv)
                item.setSelected(True)
                self.lst_grouped.addItem(item)

        self.update_next_button_state()

    def populate_metrics_checkboxes(self):
        """
        Metrik sütunlar için onay kutuları oluşturur.
        Boş sütunlar devre dışı bırakılır ve gri renkle gösterilir.
        """
        self.clear_metrics_checkboxes()

        self.main_window.selected_metrics = []  # Seçili metrikleri sıfırla

        if not self.main_window.metric_cols:
            # Metrik yoksa uyarı mesajı göster
            empty_label = QLabel("Seçilebilir metrik bulunamadı.", parent=self.metrics_content_widget)
            empty_label.setAlignment(Qt.AlignCenter)
            self.metrics_layout.addWidget(empty_label)
            self.btn_next.setEnabled(False)
            return

        for col_name in self.main_window.metric_cols:
            checkbox = QCheckBox(col_name)
            # Sütunun tamamen boş olup olmadığını kontrol et
            is_entirely_empty = self.main_window.df[col_name].dropna().empty

            if is_entirely_empty:
                checkbox.setChecked(False)
                checkbox.setEnabled(False)  # Boş sütunları pasif yap
                checkbox.setText(f"{col_name} (Boş)")
                checkbox.setStyleSheet("color: gray;")  # Gri renk
            else:
                checkbox.setChecked(True)  # Varsayılan seçili
                self.main_window.selected_metrics.append(col_name)

            checkbox.stateChanged.connect(self.on_metric_checkbox_changed)
            self.metrics_layout.addWidget(checkbox)

        self.metrics_layout.addStretch(1)  # Checkboxları üstte topla
        self.update_next_button_state()

    def clear_metrics_checkboxes(self):
        """Metrik onay kutularını temizler ve layout'tan kaldırır."""
        while self.metrics_layout.count():
            item = self.metrics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
            elif isinstance(item, QSpacerItem):
                self.metrics_layout.removeItem(item)

    def on_metric_checkbox_changed(self, state):
        """Metrik onay kutusunun durumu değiştiğinde seçili metrik listesini günceller."""
        sender_checkbox = self.sender()
        metric_name = sender_checkbox.text().replace(" (Boş)", "")  # "(Boş)" ifadesini temizle

        if state == Qt.Checked:
            if metric_name not in self.main_window.selected_metrics:
                self.main_window.selected_metrics.append(metric_name)
        else:
            if metric_name in self.main_window.selected_metrics:
                self.main_window.selected_metrics.remove(metric_name)

        self.update_next_button_state()

    def update_next_button_state(self):
        """
        İleri butonunun aktifliğini belirler.
        En az bir gruplanan öğe ve en az bir metrik seçilmiş olmalı.
        """
        is_grouped_selected = bool(self.lst_grouped.selectedItems())
        is_metric_selected = bool(self.main_window.selected_metrics)
        self.btn_next.setEnabled(is_grouped_selected and is_metric_selected)

    def go_next(self) -> None:
        """
        İleri butonuna basıldığında seçilen değerleri ana pencereye kaydeder
        ve sonraki sayfaya geçiş yapar.
        """
        self.main_window.grouped_values = [i.text() for i in self.lst_grouped.selectedItems()]
        self.main_window.selected_grouping_val = self.cmb_grouping.currentText()
        if not self.main_window.grouped_values or not self.main_window.selected_metrics:
            QMessageBox.warning(self, "Seçim Eksik", "Lütfen en az bir gruplanan değişken ve bir metrik seçin.")
            return
        self.main_window.goto_page(2)

    def _update_sheet_selection(self) -> None:
        """
        Excel dosyasındaki sayfa listesini günceller, geçerli sayfayı seçer,
        ve ilgili verileri yükler. Sinyaller geçici olarak engellenir.
        """
        self.cmb_sheet.blockSignals(True)

        self.cmb_sheet.clear()
        if self.main_window.available_sheets:
            # "KAPLAMA-OEE" sayfası dışındaki sayfaları filtrele
            daily_graph_sheets = [sheet for sheet in self.main_window.available_sheets if sheet != "KAPLAMA-OEE"]
            if daily_graph_sheets:
                self.cmb_sheet.addItems(daily_graph_sheets)
                self.cmb_sheet.setEnabled(True)
                # Varsayılan seçimi "SMD-OEE" yap, yoksa ilk sayfayı seç
                if "SMD-OEE" in daily_graph_sheets:
                    self.cmb_sheet.setCurrentText("SMD-OEE")
                else:
                    self.cmb_sheet.setCurrentText(daily_graph_sheets[0])
            else:
                self.cmb_sheet.setEnabled(False)
                self.main_window.selected_sheet = None
                QMessageBox.warning(self, "Uyarı", "Günlük grafikler için uygun sayfa bulunamadı.")
                self.cmb_sheet.blockSignals(False)
                self.main_window.goto_page(0)
                return
        else:
            self.cmb_sheet.setEnabled(False)
            self.main_window.selected_sheet = None
            QMessageBox.warning(self, "Uyarı", "Seçilen Excel dosyasında uygun sayfa bulunamadı.")
            self.cmb_sheet.blockSignals(False)
            self.main_window.goto_page(0)
            return

        self.cmb_sheet.blockSignals(False)

        # Seçili sayfayı kaydet, Excel verisini yükle ve seçim alanlarını doldur
        self.main_window.selected_sheet = self.cmb_sheet.currentText()
        self.main_window.load_excel()
        self._populate_data_selection_fields()

    def refresh(self) -> None:
        """Sayfa görüntülendiğinde çağrılır ve verileri yeniler."""
        self._update_sheet_selection()

    def on_sheet_selected(self) -> None:
        """
        Sayfa seçimi değiştiğinde çağrılır.
        Ana pencereye seçimi bildirir ve verileri yeniden yükler.
        """
        self._update_sheet_selection()
