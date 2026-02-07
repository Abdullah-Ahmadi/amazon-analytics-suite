"""
User interface handling module with GUI dialogs.
"""

import sys
import os
from pathlib import Path
from datetime import datetime
import logging

# Try to import PyQt5, fall back to console mode
try:
    from PyQt5.QtWidgets import (QApplication, QMessageBox, QProgressDialog,
                                 QFileDialog, QMainWindow, QLabel, QPushButton,
                                 QVBoxLayout, QWidget)
    from PyQt5.QtCore import Qt, QTimer
    QT_AVAILABLE = True
except ImportError:
    QT_AVAILABLE = False
    print("PyQt5 not available. Running in console mode.")

class UIManager:
    """Manages user interface and interactions."""

    def __init__(self, config, logger):
        self.config = config
        self.logger = logger
        self.app = None

        if QT_AVAILABLE:
            self.app = QApplication(sys.argv)
            self.window = None

    def run(self):
        """Run the user interface."""
        if QT_AVAILABLE:
            self._run_gui()
        else:
            self._run_console()

    def _run_gui(self):
        """Run with PyQt5 GUI."""
        self.logger.info("Starting GUI mode...")

        try:
            # Create main window
            self.window = QMainWindow()
            self.window.setWindowTitle("Amazon Analytics Suite")
            self.window.setGeometry(100, 100, 400, 200)

            # Create central widget
            central_widget = QWidget()
            self.window.setCentralWidget(central_widget)

            # Create layout
            layout = QVBoxLayout()

            # Title
            title = QLabel("📊 Amazon Analytics Suite")
            title.setAlignment(Qt.AlignCenter)
            title.setStyleSheet("font-size: 18px; font-weight: bold; margin: 20px;")
            layout.addWidget(title)

            # Subtitle
            subtitle = QLabel("Process Amazon CSV files and generate insights dashboard")
            subtitle.setAlignment(Qt.AlignCenter)
            subtitle.setStyleSheet("font-size: 12px; color: #666; margin-bottom: 30px;")
            layout.addWidget(subtitle)

            # Process button
            process_btn = QPushButton("Start Processing")
            process_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4472C4;
                    color: white;
                    font-weight: bold;
                    padding: 10px;
                    border-radius: 5px;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #2F5496;
                }
            """)
            process_btn.clicked.connect(self._process_files_gui)
            layout.addWidget(process_btn)

            # Status label
            self.status_label = QLabel("Ready to process files in current directory")
            self.status_label.setAlignment(Qt.AlignCenter)
            self.status_label.setStyleSheet("font-size: 11px; color: #666; margin-top: 20px;")
            layout.addWidget(self.status_label)

            central_widget.setLayout(layout)

            # Show window
            self.window.show()
            sys.exit(self.app.exec_())

        except Exception as e:
            self.logger.error(f"GUI error: {e}")
            self._run_console()

    def _run_console(self):
        """Run in console mode."""
        self.logger.info("Running in console mode...")

        print("\n" + "="*60)
        print("📊 AMAZON ANALYTICS SUITE")
        print("="*60)
        print("\nLooking for CSV files in current directory...")

        try:
            # Import here to avoid circular imports
            from data_loader import DataLoader
            from analyzer import DataAnalyzer
            from excel_writer import ExcelReportWriter

            # Initialize components
            loader = DataLoader(self.logger)
            analyzer = DataAnalyzer(self.config, self.logger)

            # Discover files
            files = loader.discover_files()

            # Check if we have any files
            total_files = sum(len(file_list) for file_list in files.values())
            if total_files == 0:
                print("❌ No CSV files found in current directory.")
                print("Please place your Amazon CSV files in the same directory as this script.")
                input("\nPress Enter to exit...")
                return

            print(f"\n✅ Found {total_files} CSV files:")
            for file_type, file_list in files.items():
                if file_list:
                    print(f"  • {file_type.title()}: {len(file_list)} files")

            # Process files
            print("\n" + "="*60)
            print("Processing files...")
            print("="*60)

            # Load data
            data_dict = {}
            data_dict['sales'] = loader.load_sales_data(files['sales']) if files['sales'] else None
            data_dict['inventory'] = loader.load_inventory_data(files['inventory']) if files['inventory'] else None
            data_dict['advertising'] = loader.load_advertising_data(files['advertising']) if files['advertising'] else None
            data_dict['reviews'] = loader.load_reviews_data(files['reviews']) if files['reviews'] else None

            # Analyze data
            analyzer.analyze_all(data_dict)

            # Generate Excel report
            print("\n" + "="*60)
            print("Generating Excel report...")
            print("="*60)

            output_dir = Path(self.config.OUTPUT_DIR)
            output_dir.mkdir(exist_ok=True)
            output_path = output_dir / self.config.output_filename

            writer = ExcelReportWriter(self.config, self.logger)
            success = writer.create_report(data_dict, analyzer, output_path)

            if success:
                print(f"\n✅ SUCCESS!")
                print(f"📁 Report saved to: {output_path}")
                print(f"📊 Sheets generated:")
                print(f"  • Executive Dashboard")
                print(f"  • Sales Analysis")
                print(f"  • Inventory Health")
                print(f"  • Advertising ROI")
                print(f"  • Customer Reviews")
                print(f"  • Actionable Alerts")
                print(f"  • Raw Data Sheets")

                # Show summary
                kpis = analyzer.get_kpi_summary()
                print(f"\n📈 KEY METRICS:")
                print(f"  • Total Revenue: ${kpis.get('total_revenue', 0):,.2f}")
                print(f"  • Total Products: {kpis.get('total_products', 0):,}")
                print(f"  • Low Stock Items: {kpis.get('low_stock_products', 0):,}")
                print(f"  • Critical Alerts: {len([a for a in analyzer.alerts if a['level'] == 'critical'])}")

            else:
                print("\n❌ Failed to generate report. Check log file for details.")

            print("\n" + "="*60)
            input("Press Enter to exit...")

        except Exception as e:
            print(f"\n❌ ERROR: {e}")
            self.logger.error(f"Console mode error: {e}")
            input("\nPress Enter to exit...")

    def _process_files_gui(self):
        """Process files from GUI."""
        try:
            # Import here to avoid circular imports
            from data_loader import DataLoader
            from analyzer import DataAnalyzer
            from excel_writer import ExcelReportWriter

            # Create progress dialog
            progress = QProgressDialog("Processing files...", "Cancel", 0, 100, self.window)
            progress.setWindowTitle("Amazon Analytics Suite")
            progress.setWindowModality(Qt.WindowModal)
            progress.setValue(0)

            # Initialize components
            loader = DataLoader(self.logger)
            analyzer = DataAnalyzer(self.config, self.logger)

            # Step 1: Discover files
            self.status_label.setText("Discovering CSV files...")
            progress.setValue(10)
            QApplication.processEvents()

            files = loader.discover_files()
            total_files = sum(len(file_list) for file_list in files.values())

            if total_files == 0:
                QMessageBox.warning(self.window, "No Files Found",
                                  "No CSV files found in current directory.\n\n"
                                  "Please place your Amazon CSV files in the same directory as this script.")
                return

            # Step 2: Load data
            self.status_label.setText("Loading data...")
            progress.setValue(30)
            QApplication.processEvents()

            data_dict = {}
            data_dict['sales'] = loader.load_sales_data(files['sales']) if files['sales'] else None
            data_dict['inventory'] = loader.load_inventory_data(files['inventory']) if files['inventory'] else None
            data_dict['advertising'] = loader.load_advertising_data(files['advertising']) if files['advertising'] else None
            data_dict['reviews'] = loader.load_reviews_data(files['reviews']) if files['reviews'] else None

            # Step 3: Analyze data
            self.status_label.setText("Analyzing data...")
            progress.setValue(60)
            QApplication.processEvents()

            analyzer.analyze_all(data_dict)

            # Step 4: Generate report
            self.status_label.setText("Generating Excel report...")
            progress.setValue(80)
            QApplication.processEvents()

            output_dir = Path(self.config.OUTPUT_DIR)
            output_dir.mkdir(exist_ok=True)
            output_path = output_dir / self.config.output_filename

            writer = ExcelReportWriter(self.config, self.logger)
            success = writer.create_report(data_dict, analyzer, output_path)

            # Step 5: Complete
            progress.setValue(100)
            QApplication.processEvents()

            if success:
                self.status_label.setText(f"Report generated: {output_path.name}")

                # Show success message
                msg = QMessageBox(self.window)
                msg.setIcon(QMessageBox.Information)
                msg.setWindowTitle("Success!")
                msg.setText("✅ Report Generated Successfully!")
                msg.setInformativeText(
                    f"File: {output_path.name}\n"
                    f"Location: {output_path.parent}\n\n"
                    f"Total Alerts: {len(analyzer.alerts)}\n"
                    f"Total Products: {analyzer.get_kpi_summary().get('total_products', 0):,}"
                )
                msg.setStandardButtons(QMessageBox.Ok)
                msg.exec_()

            else:
                self.status_label.setText("Failed to generate report")
                QMessageBox.critical(self.window, "Error",
                                   "Failed to generate report. Check log file for details.")

        except Exception as e:
            self.logger.error(f"GUI processing error: {e}")
            QMessageBox.critical(self.window, "Error",
                               f"An error occurred:\n\n{str(e)}")

    def show_info(self, title, message, icon="info"):
        """Show information dialog."""
        if QT_AVAILABLE:
            icon_map = {
                "info": QMessageBox.Information,
                "warning": QMessageBox.Warning,
                "error": QMessageBox.Critical,
                "question": QMessageBox.Question
            }

            msg = QMessageBox(self.window if self.window else None)
            msg.setIcon(icon_map.get(icon, QMessageBox.Information))
            msg.setWindowTitle(title)
            msg.setText(message)
            msg.exec_()
        else:
            print(f"\n{title.upper()}: {message}")

    def show_progress(self, message, value, maximum=100):
        """Show progress update."""
        if QT_AVAILABLE:
            # Could implement with QProgressDialog
            pass
        else:
            print(f"{message}... {value}/{maximum}")
