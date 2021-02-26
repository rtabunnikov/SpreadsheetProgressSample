using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using DevExpress.Office.Services;
using DevExpress.Office.Services.Implementation;
using DevExpress.Services;
using DevExpress.XtraSpreadsheet.Export;

namespace SpreadsheetProgressSample {
    public partial class Form1 : Form, IProgressIndicationService {
        CancellationTokenSource cancellationTokenSource;
        ICancellationTokenProvider savedCancellationTokenProvider;

        public Form1() {
            InitializeComponent();
            spreadsheetControl1.ReplaceService<IProgressIndicationService>(this);
            //spreadsheetControl1.Options.Compatibility.EnableLegacyPdfExport = true;
        }

        public void Begin(string displayName, int minProgress, int maxProgress, int currentProgress) {
            cancellationTokenSource = new CancellationTokenSource();
            savedCancellationTokenProvider = spreadsheetControl1.ReplaceService<ICancellationTokenProvider>(new CancellationTokenProvider(cancellationTokenSource.Token));
            repositoryItemProgressBar1.Minimum = minProgress;
            repositoryItemProgressBar1.Maximum = maxProgress;
            barProgress.Caption = displayName;
            barProgress.EditValue = currentProgress;
            butCancel.Enabled = true;
        }

        public void End() {
            spreadsheetControl1.ReplaceService(savedCancellationTokenProvider);
            cancellationTokenSource?.Dispose();
            cancellationTokenSource = null;
            barProgress.Caption = "";
            barProgress.EditValue = 0;
            butCancel.Enabled = false;
        }

        public void SetProgress(int currentProgress) {
            barProgress.EditValue = currentProgress;
            Application.DoEvents();
        }

        private void btnExportToHtml_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            try {
                spreadsheetControl1.Document.ExportToHtml("c:\\test\\_test.html",
                    new HtmlDocumentExporterOptions() { SheetIndex = 0, EmbedImages = true, ExportImages = true });
            }
            catch (OperationCanceledException) {
                End();
            }
        }

        private void btnExportToPDF_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            try {
                spreadsheetControl1.Document.ExportToPdf("c:\\test\\_test.pdf");
            }
            catch (OperationCanceledException) {
                End();
            }
        }

        private void butCancel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            cancellationTokenSource?.Cancel();
        }

        private void spreadsheetControl1_UnhandledException(object sender, DevExpress.XtraSpreadsheet.SpreadsheetUnhandledExceptionEventArgs e) {
            if (e.Exception is OperationCanceledException) {
                e.Handled = true;
                End();
            }
        }
    }
}
