import logging
from pathlib import Path
import re
from typing import List, Tuple, Any, Union, Dict
from utils.helpers import seconds_from_timedelta
import pandas as pd

from PyQt5.QtCore import QThread, pyqtSignal
from utils.helpers import excel_col_to_index


class MonthlyGraphWorker(QThread):
    """
    Aylık grafik oluşturma için arka planda çalışan iş parçacığı sınıfı.

    Args:
        excel_path (Path): İşlenecek Excel dosyasının yolu.
        current_df (pd.DataFrame): Ana pencereden gelen ve işlenecek mevcut DataFrame.
        graph_mode (str): Grafik modu ("hat" veya "page").
        graph_type (str): Grafik türü ("OEE Grafikleri", "Dizgi Onay Dağılım Grafiği", "Dizgi Duruş Grafiği").
        prev_year_oee (float | None): Önceki yıl OEE değeri (isteğe bağlı).
        prev_month_oee (float | None): Önceki ay OEE değeri (isteğe bağlı).
        main_window (MainWindow): Ana pencereye referans, gerekli özelliklere erişim için.
    """
    finished = pyqtSignal(list, object, object)  # Grafik verisi, önceki yıl ve önceki ay OEE iletim sinyali
    progress = pyqtSignal(int)  # İlerleme yüzdesi sinyali
    error = pyqtSignal(str)  # Hata mesajı sinyali

    def __init__(self, excel_path: Path, current_df: pd.DataFrame, graph_mode: str, graph_type: str,
                 prev_year_oee: float | None, prev_month_oee: float | None, main_window: "MainWindow"):
        super().__init__()
        self.excel_path = excel_path
        self.current_df = current_df  # Ana pencereden gelen DataFrame (genellikle SMD-OEE)
        self.graph_mode = graph_mode
        self.graph_type = graph_type
        self.prev_year_oee = prev_year_oee
        self.prev_month_oee = prev_month_oee
        self.main_window = main_window  # Ana pencere referansı

    def run(self):
        """
        İş parçacığı çalışmaya başladığında çağrılır.
        Excel dosyasından veya mevcut DataFrame'den grafik verilerini oluşturur,
        işlemin ilerlemesini ve varsa hataları ilgili sinyallerle bildirir.
        """
        try:
            figures_data: List[Tuple[str, Union[List[dict[str, Any]], Dict[str, Any]]]] = []

            # Grafik modu "hat" ise hat bazlı verileri işle
            if self.graph_mode == "hat":
                df_to_process = self.current_df.copy()

                # Ana pencereden sütun isimlerini al
                grouping_col_name = self.main_window.grouping_col_name
                grouped_col_name = self.main_window.grouped_col_name
                oee_col_name = self.main_window.oee_col_name

                # Sütunları dahili tutarlılık için yeniden adlandır
                col_mapping = {}
                if grouping_col_name in df_to_process.columns:
                    col_mapping[grouping_col_name] = 'Tarih'
                if grouped_col_name in df_to_process.columns:
                    col_mapping[grouped_col_name] = 'U_Agaci_Sev'
                if oee_col_name and oee_col_name in df_to_process.columns:
                    col_mapping[oee_col_name] = 'OEE_Degeri'

                # Sütun adları uygun değilse hata sinyali gönder
                if col_mapping:
                    df_to_process.rename(columns=col_mapping, inplace=True)
                else:
                    self.error.emit("Gerekli sütunlar Excel dosyasında bulunamadı veya adlandırılamadı.")
                    return

                # 'Tarih' sütununu datetime türüne dönüştür ve geçersiz kayıtları kaldır
                if 'Tarih' in df_to_process.columns:
                    df_to_process['Tarih'] = pd.to_datetime(df_to_process['Tarih'], errors='coerce')
                    df_to_process.dropna(subset=['Tarih'], inplace=True)
                else:
                    self.error.emit("'Tarih' sütunu bulunamadı.")
                    return

                # OEE grafikleri için 'OEE_Degeri' sütununu float'a dönüştür
                if self.graph_type == "OEE Grafikleri":
                    if 'OEE_Degeri' in df_to_process.columns:
                        df_to_process['OEE_Degeri'] = pd.to_numeric(
                            df_to_process['OEE_Degeri'].astype(str).replace('%', '').replace(',', '.'),
                            errors='coerce'
                        )
                        df_to_process['OEE_Degeri'] = df_to_process['OEE_Degeri'].fillna(0.0)

                        # Eğer OEE değerleri % formatındaysa 100'e böl
                        if not df_to_process['OEE_Degeri'].empty and df_to_process['OEE_Degeri'].max() > 1.0:
                            df_to_process['OEE_Degeri'] = df_to_process['OEE_Degeri'] / 100.0
                    else:
                        self.error.emit("'OEE_Degeri' sütunu bulunamadı.")
                        return

                # Dizgi Onay Dağılım Grafiği için sütun indeksi (T sütunu)
                dizgi_onay_col_index = excel_col_to_index('T')
                dizgi_onay_col_name = self.current_df.columns[dizgi_onay_col_index] if dizgi_onay_col_index < len(
                    self.current_df.columns) else None

                # Dizgi Duruş Grafiği için metrik sütunları (H'den BD'ye kadar)
                dizgi_durusu_metric_cols = []
                start_col_index = excel_col_to_index('H')
                end_col_index = excel_col_to_index('BD')
                for i in range(start_col_index, end_col_index + 1):
                    if i < len(self.current_df.columns):
                        col_name = self.current_df.columns[i]
                        dizgi_durusu_metric_cols.append(col_name)

                # Dizgi Onay Dağılım Grafiği için süreci hazırla
                if self.graph_type == "Dizgi Onay Dağılım Grafiği":
                    if not dizgi_onay_col_name or dizgi_onay_col_name not in df_to_process.columns:
                        self.error.emit(f"'{dizgi_onay_col_name}' (Dizgi Onay) sütunu bulunamadı veya geçersiz.")
                        return
                    # Süreyi saniyeye çevir
                    df_to_process[dizgi_onay_col_name] = seconds_from_timedelta(df_to_process[dizgi_onay_col_name])

                # Dizgi Duruş Grafiği için metrik sütunları süreye çevir
                elif self.graph_type == "Dizgi Duruş Grafiği":
                    if not dizgi_durusu_metric_cols:
                        self.error.emit("Dizgi Duruş Grafiği için metrik sütunları bulunamadı.")
                        return
                    for col in dizgi_durusu_metric_cols:
                        if col in df_to_process.columns:
                            df_to_process[col] = seconds_from_timedelta(df_to_process[col])

                # 'Group_Key' sütunu oluştur: "HAT" ile başlayan ve formatlanmış stringler
                def extract_group_key(s):
                    s = str(s).upper()
                    match = re.search(r'HAT(\d+)', s)
                    if match:
                        hat_number = match.group(1)
                        return f"HAT-{hat_number}"
                    return None

                # 'U_Agaci_Sev' sütunu varsa grup anahtarlarını çıkar
                if 'U_Agaci_Sev' in df_to_process.columns:
                    df_to_process['Group_Key'] = df_to_process['U_Agaci_Sev'].apply(extract_group_key)
                    df_to_process.dropna(subset=['Group_Key'], inplace=True)
                else:
                    self.error.emit("'U_Agaci_Sev' sütunu bulunamadı.")
                    return

                # Mevcut hatlar filtrelenir ve sıralanır
                unique_hats = sorted(df_to_process['Group_Key'].unique())
                target_hat_patterns = {"HAT-1", "HAT-2", "HAT-3", "HAT-4"}
                filtered_hats = [hat for hat in unique_hats if hat in target_hat_patterns]
                unique_hats = sorted(filtered_hats)
                total_items = len(unique_hats)

                # Hat verisi yoksa hata mesajı gönder
                if not unique_hats and self.graph_type != "Dizgi Duruş Grafiği":
                    self.error.emit(
                        "Grafik oluşturmak için hat verisi bulunamadı. Lütfen Excel dosyasının 'HAT-1', 'HAT-2', 'HAT-3' veya 'HAT-4' için veri içerdiğinden emin olun.")
                    return

                # Dizgi Duruş Grafiği için Pareto analizi yapılır
                if self.graph_type == "Dizgi Duruş Grafiği":
                    all_metrics_sum = df_to_process[dizgi_durusu_metric_cols].sum()
                    total_sum_of_all_metrics = all_metrics_sum.sum()
                    metric_sums = all_metrics_sum[all_metrics_sum > 0].sort_values(ascending=False)
                    cumulative_sum_for_line = metric_sums.cumsum()
                    cumulative_percentage_for_line = (cumulative_sum_for_line / total_sum_of_all_metrics) * 100

                    pareto_metrics_to_plot = pd.Series(dtype=float)
                    current_cumulative_percent = 0.0
                    for idx, (metric_name, value) in enumerate(metric_sums.items()):
                        percent_of_total = (value / total_sum_of_all_metrics) * 100
                        current_cumulative_percent += percent_of_total
                        pareto_metrics_to_plot[metric_name] = value
                        if current_cumulative_percent >= 80:
                            if current_cumulative_percent - percent_of_total >= 80 and (
                                    current_cumulative_percent - 80) > 10:
                                pareto_metrics_to_plot = pareto_metrics_to_plot.iloc[:-1]
                            break
                    if pareto_metrics_to_plot.empty and not metric_sums.empty:
                        pareto_metrics_to_plot = metric_sums.head(1)

                    # Pareto grafiği verileri figures_data'ya eklenir
                    figures_data.append(("Genel Dizgi Duruş", {
                        "metrics": pareto_metrics_to_plot.to_dict(),
                        "total_overall_sum": total_sum_of_all_metrics,
                        "cumulative_percentages": cumulative_percentage_for_line[pareto_metrics_to_plot.index].to_dict()
                    }))
                    self.progress.emit(100)

                # Diğer grafik türleri için hat bazlı verileri işle
                else:
                    for i, selected_hat in enumerate(unique_hats):
                        df_smd_oee_filtered_by_hat = df_to_process[df_to_process['Group_Key'] == selected_hat].copy()
                        if df_smd_oee_filtered_by_hat.empty:
                            self.progress.emit(int((i + 1) / total_items * 100))
                            continue

                        # OEE grafikleri için günlük ortalamaları hesapla
                        if self.graph_type == "OEE Grafikleri":
                            grouped_oee = df_smd_oee_filtered_by_hat.groupby(pd.Grouper(key='Tarih', freq='D'))[
                                'OEE_Degeri'].mean().reset_index()
                            grouped_oee.dropna(subset=['OEE_Degeri'], inplace=True)
                            figures_data.append((selected_hat, grouped_oee.to_dict('records')))

                        # Dizgi Onay Dağılım Grafiği için verileri hazırla
                        elif self.graph_type == "Dizgi Onay Dağılım Grafiği":
                            current_hat_onay_sum = df_smd_oee_filtered_by_hat[dizgi_onay_col_name].sum()
                            other_hats_df = df_to_process[df_to_process['Group_Key'] != selected_hat].copy()
                            other_hats_onay_sum = other_hats_df[dizgi_onay_col_name].sum()
                            total_onay_sum = current_hat_onay_sum + other_hats_onay_sum
                            if total_onay_sum > 0:
                                figures_data.append((selected_hat, [
                                    {"label": selected_hat, "value": current_hat_onay_sum},
                                    {"label": "DİĞER HATLAR", "value": other_hats_onay_sum}
                                ]))
                        # İlerleme sinyali gönder
                        self.progress.emit(int((i + 1) / total_items * 100))

            # Grafik modu "page" ise ve grafik türü OEE ise sayfa bazlı işlemler
            elif self.graph_mode == "page" and self.graph_type == "OEE Grafikleri":
                # İşlenecek sayfalar ve OEE sütun harfleri tanımlanır
                sheets_to_process_info = [
                    ("DALGA_LEHİM", "BP"),
                    ("ROBOT", "BG"),
                    ("KAPLAMA-OEE", "BG")
                ]

                # Excel dosyasında mevcut olan sayfalar filtrelenir
                available_sheets_for_page_mode = [
                    (sheet_name, oee_col) for sheet_name, oee_col in sheets_to_process_info
                    if sheet_name in self.main_window.available_sheets
                ]

                total_items = len(available_sheets_for_page_mode)
                if not available_sheets_for_page_mode:
                    self.error.emit(
                        "Sayfa grafikleri için işlenecek uygun sayfa bulunamadı (DALGA_LEHİM, ROBOT, KAPLAMA-OEE).")
                    return

                # Her sayfa için veri işleme
                for i, (sheet_name, oee_col_letter) in enumerate(available_sheets_for_page_mode):
                    logging.info(
                        f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfası için OEE grafiği oluşturuluyor...")

                    try:
                        # Sayfa verisini oku
                        sheet_df = pd.read_excel(self.excel_path, sheet_name=sheet_name, header=0)
                        sheet_df.columns = sheet_df.columns.astype(str)  # Sütun isimleri string olarak ayarlanır
                    except Exception as e:
                        logging.warning(f"'{sheet_name}' sayfası yüklenirken hata oluştu: {e}. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    # Tarih sütununu al (A sütunu)
                    tarih_col_name = sheet_df.columns[excel_col_to_index('A')] if excel_col_to_index('A') < len(
                        sheet_df.columns) else None
                    if not tarih_col_name or tarih_col_name not in sheet_df.columns:
                        logging.warning(
                            f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfasında 'A' sütunu (Tarih) bulunamadı. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    # OEE sütununu al
                    current_oee_col_index = excel_col_to_index(oee_col_letter)
                    current_oee_col_name = sheet_df.columns[current_oee_col_index] if current_oee_col_index < len(
                        sheet_df.columns) else None

                    if not current_oee_col_name or current_oee_col_name not in sheet_df.columns:
                        logging.warning(
                            f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfası için '{oee_col_letter}' ({current_oee_col_name}) sütunu bulunamadı veya geçersiz. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    # Tarih sütununu datetime türüne çevir ve geçersizleri temizle
                    sheet_df['Tarih'] = pd.to_datetime(sheet_df[tarih_col_name], errors='coerce')
                    sheet_df.dropna(subset=['Tarih'], inplace=True)

                    if sheet_df.empty:
                        logging.warning(
                            f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfası için tarih verisi bulunamadı. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    # OEE sütununu sayısal değere dönüştür
                    sheet_df['OEE_Degeri_Processed'] = pd.to_numeric(
                        sheet_df[current_oee_col_name].astype(str).replace('%', '').replace(',', '.'),
                        errors='coerce'
                    )
                    sheet_df['OEE_Degeri_Processed'] = sheet_df['OEE_Degeri_Processed'].fillna(0.0)
                    # % formatında ise 100'e böl
                    if not sheet_df['OEE_Degeri_Processed'].empty and sheet_df['OEE_Degeri_Processed'].max() > 1.0:
                        sheet_df['OEE_Degeri_Processed'] = sheet_df['OEE_Degeri_Processed'] / 100.0

                    # Tarihe göre grupla ve günlük ortalama OEE değerini hesapla
                    grouped_oee = sheet_df.groupby(pd.Grouper(key='Tarih', freq='D'))[
                        'OEE_Degeri_Processed'].mean().reset_index()
                    grouped_oee.dropna(subset=['OEE_Degeri_Processed'], inplace=True)

                    # Yarı değer çizgisi hesapla (ek grafik için kullanılabilir)
                    grouped_oee['OEE_Degeri_Half'] = grouped_oee['OEE_Degeri_Processed'] / 2

                    if grouped_oee.empty:
                        logging.warning(
                            f"MonthlyGraphWorker (Page Mode): '{sheet_name}' sayfası için işlenecek OEE verisi bulunamadı. Atlanıyor.")
                        self.progress.emit(int((i + 1) / total_items * 100))
                        continue

                    # İşlenmiş veriyi figures_data listesine ekle
                    figures_data.append((sheet_name, grouped_oee.to_dict('records')))
                    self.progress.emit(int((i + 1) / total_items * 100))

            # İşlem tamamlandığında sonuçları ve önceki OEE değerlerini gönder
            self.finished.emit(figures_data, self.prev_year_oee, self.prev_month_oee)

        except Exception as exc:
            logging.exception("MonthlyGraphWorker hatası oluştu.")
            # Oluşan hata mesajını dışarı ilet
            self.error.emit(f"Aylık grafik oluşturulurken bir hata oluştu: {str(exc)}")
