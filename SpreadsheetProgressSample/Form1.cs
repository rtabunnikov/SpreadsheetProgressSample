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
            savedCancellationTokenProvider = spreadsheetControl1.ReplaceService<ICancellationTokenProvider>(new CancellationTokenProvider(cancellationTokenSource.Token));
            splashScreenManager1.ShowWaitForm();
            splashScreenManager1.SetWaitFormCaption(displayName);
            splashScreenManager1.SetWaitFormDescription($"{currentProgress}%");
            splashScreenManager1.SendCommand(WaitForm1.WaitFormCommand.SetCancellationTokenSource, cancellationTokenSource);
            Application.DoEvents();
        }

        void IProgressIndicationService.End() {
            splashScreenManager1.CloseWaitForm();
            spreadsheetControl1.ReplaceService(savedCancellationTokenProvider);
            spreadsheetControl1.UpdateCommandUI();
            cancellationTokenSource?.Dispose();
            cancellationTokenSource = null;
        }

        void IProgressIndicationService.SetProgress(int currentProgress) {
            splashScreenManager1.SetWaitFormDescription($"{currentProgress}%");
        }

        void spreadsheetControl1_UnhandledException(object sender, DevExpress.XtraSpreadsheet.SpreadsheetUnhandledExceptionEventArgs e) {
            if (e.Exception is OperationCanceledException)
                e.Handled = true;
        }
    }
}
