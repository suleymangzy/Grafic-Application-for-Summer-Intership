from typing import List, Any
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors
import numpy as np

class GraphPlotter:
    """Matplotlib grafikleri (donut ve çubuk) oluşturmak için yardımcı sınıf."""

    @staticmethod
    def create_donut_chart(
            ax: plt.Axes,  # Donut grafiği çizilecek eksen
            sorted_metrics_series: pd.Series,  # Sıralı metrik verileri (isim + süre)
            oee_display_value: str,  # OEE değeri (gösterim için)
            chart_colors: List[Any],  # Dilim renkleri
            fig: plt.Figure  # Matplotlib figürü (etiketleri dışa basmak için)
    ) -> None:
        """Donut (halka) grafik oluşturur ve üzerine OEE değeri ile en çok 3 duruş bilgisini ekler."""

        # Donut grafik oluşturuluyor (width < 1 çemberi içe boşluklu yapar)
        wedges, texts = ax.pie(
            sorted_metrics_series,
            autopct=None,
            startangle=90,  # Saat 12 yönünden başlasın
            wedgeprops=dict(width=0.4, edgecolor='w'),  # Dilimler arası beyaz çizgi
            colors=chart_colors[:len(sorted_metrics_series)]  # Renkler
        )

        # OEE değeri merkezde gösterilir
        oee_text_to_display = f"OEE\n{oee_display_value}" if oee_display_value else "OEE\nVeri Yok"
        ax.text(0, 0, oee_text_to_display,
                horizontalalignment='center', verticalalignment='center',
                fontsize=28, fontweight='bold', color='black')  # Merkezde büyük fontlu metin

        # Etiketlerin başlayacağı y koordinatı hesaplanır (sayfa boyutuna göre dinamik)
        label_y_start = 0.25 + (30 / (fig.get_size_inches()[1] * fig.dpi))
        label_line_height = 0.05  # Her bir satır arası boşluk

        top_3_metrics = sorted_metrics_series.head(3)  # En uzun süren 3 duruş
        top_3_colors = chart_colors[:len(top_3_metrics)]  # Onlara ait renkler

        for i, (metric_name, metric_value) in enumerate(top_3_metrics.items()):
            # Süreyi saat:dakika:saniye formatında hesapla
            duration_hours = int(metric_value // 3600)
            duration_minutes = int((metric_value % 3600) // 60)
            duration_seconds = int(metric_value % 60)
            # Etiket metni
            label_text = (
                f"{i + 1}. {metric_name}; "
                f"{duration_hours:02d}:{duration_minutes:02d}:{duration_seconds:02d}; "
                f"{metric_value / sorted_metrics_series.sum() * 100:.0f}%"
            )
            y_pos = label_y_start - (i * label_line_height)  # Her satır aşağıya kayar

            # Kutu özellikleri (renkli arka plan)
            bbox_props = dict(boxstyle="round,pad=0.3", fc=top_3_colors[i], ec=top_3_colors[i], lw=0.5)

            # Arka plan rengine göre yazı rengini ayarla (kontrast için)
            r, g, b, _ = matplotlib.colors.to_rgba(top_3_colors[i])
            luminance = (0.299 * r + 0.587 * g + 0.114 * b)
            text_color = 'white' if luminance < 0.5 else 'black'

            # Etiketi figürün sol üstüne yerleştir
            fig.text(0.005,  # Sol kenara yakın (konumu değiştirmiyoruz)
                     y_pos,
                     label_text,
                     horizontalalignment='left', verticalalignment='top',
                     fontsize=12,  # Font büyüklüğü
                     bbox=bbox_props,  # Arka plan kutusu
                     transform=fig.transFigure,
                     color=text_color)

        ax.set_title("")  # Başlık kullanılmıyor
        ax.axis("equal")  # Çemberin daire olarak çizilmesini sağlar

        # Donut grafiği sağa kaydırmak için sol boşluk artırılır (etiketlerle çakışmasın)
        fig.tight_layout(rect=[0.38, 0.1, 1, 0.95])

    @staticmethod
    def create_bar_chart(
            ax: plt.Axes,  # Çubuk grafik ekseni
            sorted_metrics_series: pd.Series,  # Metrik verileri
            oee_display_value: str,  # OEE gösterim değeri
            chart_colors: List[Any]  # Renkler
    ) -> None:
        """Yatay çubuk grafik oluşturur."""

        metrics = sorted_metrics_series.index.tolist()  # Kategori adları
        values = sorted_metrics_series.values.tolist()  # Süre değerleri (saniye)

        values_minutes = [v / 60 for v in values]  # Süreleri dakikaya çevir
        y_pos = np.arange(len(metrics))  # Y ekseni pozisyonları

        ax.barh(y_pos, values_minutes, color=chart_colors)  # Yatay çubukları çiz
        ax.set_yticks(y_pos)
        ax.set_yticklabels(metrics, fontsize=10)  # Kategori etiketlerini ayarla
        ax.invert_yaxis()  # En büyük üstte olacak şekilde ters çevir

        ax.set_xlabel("")  # X ekseni başlığı yok
        ax.set_xticks([])  # X ekseni tıklanamaz

        # Başlık: OEE gösterimi
        oee_title_text = f"OEE: {oee_display_value}" if oee_display_value else "OEE: Veri Yok"
        ax.set_title(oee_title_text, fontsize=24, fontweight='bold')

        ax.grid(False)  # Izgara çizgileri kapalı

        # Grafik çerçevesi ayarları
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['left'].set_visible(True)
        ax.spines['bottom'].set_visible(True)

        total_sum = sorted_metrics_series.sum()  # Toplam değer (yüzde hesaplamak için)
        for i, (value, metric_name) in enumerate(zip(values, metrics)):
            percentage = (value / total_sum) * 100 if total_sum > 0 else 0
            duration_hours = int(value // 3600)
            duration_minutes = int((value % 3600) // 60)
            duration_seconds = int(value % 60)
            text_label = f"{duration_hours:02d}:{duration_minutes:02d}:{duration_seconds:02d} ({percentage:.0f}%)"

            text_x_position = (value / 60) + 0.5  # Çubuğun biraz sağında yer alır
            ax.text(text_x_position, i, text_label,
                    va='center', ha='left',
                    fontsize=11, fontweight='bold',
                    color='black')

        ax.set_xlim(left=0)  # X ekseni sıfırdan başlasın
        plt.tight_layout(rect=[0.1, 0.1, 0.95, 0.9])  # Grafik kenar boşlukları
