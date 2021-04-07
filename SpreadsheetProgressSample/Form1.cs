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

namespace SpreadsheetProgressSample {
    public partial class Form1 : Form, IProgressIndicationService {
        CancellationTokenSource cancellationTokenSource;
        ICancellationTokenProvider savedCancellationTokenProvider;

        public Form1() {
            InitializeComponent();
            spreadsheetControl1.ReplaceService<IProgressIndicationService>(this);
        }

        void IProgressIndicationService.Begin(string displayName, int minProgress, int maxProgress, int currentProgress) {
            cancellationTokenSource = new CancellationTokenSource();
            savedCancellationTokenProvider = spreadsheetControl1.ReplaceService(new CancellationTokenProvider(cancellationTokenSource.Token));
            repositoryItemProgressBar1.Minimum = minProgress;
            repositoryItemProgressBar1.Maximum = maxProgress;
            barProgress.Caption = displayName;
            barProgress.EditValue = currentProgress;
            butCancel.Enabled = true;
        }

        void IProgressIndicationService.End() {
            spreadsheetControl1.ReplaceService(savedCancellationTokenProvider);
            spreadsheetControl1.UpdateCommandUI();
            cancellationTokenSource?.Dispose();
            cancellationTokenSource = null;
            barProgress.Caption = "";
            barProgress.EditValue = 0;
            butCancel.Enabled = false;
        }

        void IProgressIndicationService.SetProgress(int currentProgress) {
            barProgress.EditValue = currentProgress;
            Application.DoEvents();
        }

        void butCancel_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            cancellationTokenSource?.Cancel();
        }

        void spreadsheetControl1_UnhandledException(object sender, DevExpress.XtraSpreadsheet.SpreadsheetUnhandledExceptionEventArgs e) {
            if (e.Exception is OperationCanceledException) {
                e.Handled = true;
                ((IProgressIndicationService)this).End();
            }
        }

        void Form1_FormClosing(object sender, FormClosingEventArgs e) {
            if (cancellationTokenSource != null) {
                MessageBox.Show("Operation in progress!", Text, MessageBoxButtons.OK, MessageBoxIcon.Hand);
                e.Cancel = true;
            }
        }
    }
}
