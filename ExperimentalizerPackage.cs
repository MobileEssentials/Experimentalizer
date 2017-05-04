using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Threading;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;

namespace VisualStudio.Updater
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [Guid(PackageGuidString)]
    [ProvideAutoLoad(UIContextGuids.NoSolution)]
    public sealed class ExperimentalizerPackage : Package
    {
        const string PackageGuidString = "1c15dd58-e0b5-4215-905f-50c1113c08b1";
        static readonly string VsixExpPath = Path.Combine(
            Path.GetDirectoryName(typeof(ExperimentalizerPackage).Assembly.ManifestModule.FullyQualifiedName), "VsixExp.exe");

        Dispatcher dispatcher;

        protected override void Initialize()
        {
            base.Initialize();

            dispatcher = Dispatcher.CurrentDispatcher;

            var repo = (IVsExtensionRepository)GetService(typeof(SVsExtensionRepository));
            repo.DownloadCompleted += OnDownloadCompleted;
        }

        void OnDownloadCompleted(object sender, DownloadCompletedEventArgs e)
        {
            var factory = (IVsThreadedWaitDialogFactory)GetService(typeof(SVsThreadedWaitDialogFactory));
            IVsThreadedWaitDialog2 instance;
            factory.CreateInstance(out instance);
            var dialog = instance as IVsThreadedWaitDialog3;

            dialog.StartWaitDialogWithCallback(
                    szWaitCaption: "Experimentalizing VSIX",
                    szWaitMessage: "Preparing to process the downloaded VSIX...",
                    szProgressText: string.Empty,
                    varStatusBmpAnim: null,
                    szStatusBarText: "Experimentalizing VSIX",
                    fIsCancelable: true,
                    iDelayToShowDialog: 3,
                    fShowProgress: true, 
                    iTotalSteps: 0, 
                    iCurrentStep: 0, 
                    pCallback: new Callback());

            var info = new ProcessStartInfo(VsixExpPath, e.Payload.PackagePath)
            {
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = true,
            };

            var process = Process.Start(info);

            ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                string line;
                while (!process.HasExited && !string.IsNullOrEmpty((line = await process.StandardOutput.ReadLineAsync())))
                {
                    var cancelled = false;
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    var indexOfColon = line.LastIndexOf(':');
                    if (indexOfColon > 0)
                        line = line.Substring(indexOfColon + 1);

                    dialog.UpdateProgress("Processing downloaded VSIX...", line, line, 0, 0, false, out cancelled);
                    if (cancelled)
                    {
                        process.Kill();
                    }
                }
            });

            while (!process.HasExited)
            {
                DoEvents();
                Thread.Sleep(100);
            }

            dialog.EndWaitDialog();

            // NOTE: example of how we could uninstall before continuing. This 
            // will NOT work if the extension ID comes from a Willow component.
            // var em = (IVsExtensionManager)GetService(typeof(SVsExtensionManager));
            // var extension = em.CreateInstallableExtension(e.Payload.PackagePath);
            // if (em.IsInstalled(extension))
            // {
            //     var installed = em.GetInstalledExtension(extension.Header.Identifier);
            //     em.Uninstall(installed);
            // }
        }

        public void DoEvents()
        {
            var frame = new DispatcherFrame();
            dispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(ExitFrame), frame);
            Dispatcher.PushFrame(frame);
        }

        public object ExitFrame(object f)
        {
            ((DispatcherFrame)f).Continue = false;

            return null;
        }

        class Callback : IVsThreadedWaitDialogCallback
        {
            public void OnCanceled() { }
        }
    }
}