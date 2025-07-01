import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QMenu, QGraphicsView, QGraphicsScene
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt  # Import Qt for keyboard shortcuts


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Graphic Application")  # Corrected typo: Grafic -> Graphic
        self.setGeometry(100, 100, 800, 600)
        self.create_menu()

        # You might want to add a central widget to display graphics later
        # For now, let's just make sure the menu looks good.
        # self.scene = QGraphicsScene(self)
        # self.view = QGraphicsView(self.scene)
        # self.setCentralWidget(self.view)

    def create_menu(self):
        menubar = self.menuBar()
        # You can make the menubar native on macOS for a more integrated look
        # menubar.setNativeMenuBar(True)

        # 1. File Menu (Renamed from Dosya Seç for better common terminology)
        file_menu = menubar.addMenu("&File")  # & makes 'F' the shortcut key

        # Adding icons (you'll need actual .png or .ico files for these)
        # For demonstration, I'll assume you have some generic icons.
        # If you don't have icons, you can remove the QIcon part.

        # Word Action
        word_action = QAction(QIcon('icons/word.png'), "Open &Word File...", self)
        word_action.setShortcut("Ctrl+W")
        word_action.setStatusTip("Open a Word document")
        file_menu.addAction(word_action)

        # Excel Action
        excel_action = QAction(QIcon('icons/excel.png'), "Open &Excel File...", self)
        excel_action.setShortcut("Ctrl+E")
        excel_action.setStatusTip("Open an Excel spreadsheet")
        file_menu.addAction(excel_action)

        # PPTX Action
        pptx_action = QAction(QIcon('icons/pptx.png'), "Open &PPTX File...", self)
        pptx_action.setShortcut("Ctrl+P")
        pptx_action.setStatusTip("Open a PowerPoint presentation")
        file_menu.addAction(pptx_action)

        file_menu.addSeparator()  # Add a separator for better visual grouping

        exit_action = QAction(QIcon('icons/exit.png'), "E&xit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Exit the application")
        exit_action.triggered.connect(self.close)  # Connect to close the window
        file_menu.addAction(exit_action)

        # 2. Plot Menu (Renamed from Grafik Oluştur for better English terminology)
        plot_menu = menubar.addMenu("&Plot")

        # Matplotlib chart types
        chart_types = {
            "Line Plot": "plot",
            "Bar Chart": "bar",
            "Histogram": "hist",
            "Pie Chart": "pie",
            "Scatter Plot": "scatter",
            "Area Plot": "fill_between",
            "Box Plot": "boxplot",
            "Violin Plot": "violinplot",
            "Stem Plot": "stem",
            "Error Bar Plot": "errorbar"
        }

        for name, func_name in chart_types.items():
            action = QAction(f"Create {name}", self)
            action.setStatusTip(f"Generate a {name.lower()} using Matplotlib's '{func_name}' function")
            # You can connect these actions to specific plotting functions later
            # action.triggered.connect(lambda checked, n=name: self.create_plot(n))
            plot_menu.addAction(action)

        # 3. Data Menu (Renamed from Veri Seç for consistency)
        data_menu = menubar.addMenu("&Data")

        # X-axis selection submenu
        x_axis_menu = QMenu("Select &X-Axis", self)
        x1_action = QAction("X1 Data", self)
        x2_action = QAction("X2 Data", self)
        x_axis_menu.addAction(x1_action)
        x_axis_menu.addAction(x2_action)
        data_menu.addMenu(x_axis_menu)

        # Y-axis selection submenu
        y_axis_menu = QMenu("Select &Y-Axis", self)
        y1_action = QAction("Y1 Data", self)
        y2_action = QAction("Y2 Data", self)
        y_axis_menu.addAction(y1_action)
        y_axis_menu.addAction(y2_action)
        data_menu.addMenu(y_axis_menu)

        data_menu.addSeparator()  # Another separator

        # Example of a more specific data selection, e.g., for bar charts
        category_selection_action = QAction("Select &Category Data", self)
        category_selection_action.setStatusTip("Select data for categorical axis (e.g., for Bar Charts)")
        data_menu.addAction(category_selection_action)

        # Example for scatter plot specific options
        color_by_data_action = QAction("Color By &Data Column", self)
        color_by_data_action.setStatusTip("Choose a data column to determine point colors in Scatter Plots")
        data_menu.addAction(color_by_data_action)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())