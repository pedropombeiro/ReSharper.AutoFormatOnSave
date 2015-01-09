// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ReSharper.AutoFormatOnSavePackage.cs" company="Pedro Pombeiro">
//   2012 Pedro Pombeiro
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace ReSharper.AutoFormatOnSave
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    using EnvDTE;

    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
//// This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
//// This attribute is used to register the informations needed to show the this package
//// in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidReSharper_AutoFormatOnSavePkgString)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists)]
    public sealed class ReSharper_AutoFormatOnSavePackage : Package
    {
        #region Constants

        /// <summary>
        /// The name for the ReSharper silent code cleanup command.
        /// </summary>
        private const string ReSharperSilentCleanupCodeCommandName = "ReSharper_SilentCleanupCode";

        #endregion

        #region Static Fields

        /// <summary>
        /// The allowed file extensions for code cleanup.
        /// </summary>
        private static readonly string[] AllowedFileExtensions = new[] { ".cs", ".xaml", ".vb", ".js", ".ts", ".css", ".html", ".xml" };

        #endregion

        #region Fields

        /// <summary>
        /// The dictionary of documents which will be reformatted mapped to the timestamp when they were last reformatted.
        /// </summary>
        private readonly Dictionary<Document, DateTime> documentsToReformatDictionary = new Dictionary<Document, DateTime>();

        /// <summary>
        /// A timer which runs background checks (mainly to check whether all saves have completed).
        /// </summary>
        private readonly Timer timer = new System.Windows.Forms.Timer { Interval = 1000 };

        /// <summary>
        /// The Visual Studio build events object.
        /// </summary>
        private BuildEvents buildEvents;

        /// <summary>
        /// Positive value when the build engine is running.
        /// </summary>
        private int buildingSolution;

        /// <summary>
        /// The Visual Studio document events object.
        /// </summary>
        private DocumentEvents documentEvents;

        /// <summary>
        /// The DTE global service.
        /// </summary>
        private DTE dte;

        /// <summary>
        /// <c>true</c> if currently reformatting recently saved files.
        /// </summary>
        private bool isReformatting;

        /// <summary>
        /// The time of the last reformat.
        /// </summary>
        private DateTime lastReformat;

        /// <summary>
        /// The Visual Studio solution events object.
        /// </summary>
        private SolutionEvents solutionEvents;

        /// <summary>
        /// <c>true</c> if a solution is active.
        /// </summary>
        private bool solutionIsActive;

        #endregion

        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ReSharper_AutoFormatOnSavePackage"/> class. 
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ReSharper_AutoFormatOnSavePackage()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this));

            this.timer.Tick += this.TimerOnTick;
        }

        #endregion

        #region Methods

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                this.DisconnectFromVsEvents();

                this.timer.Tick -= this.TimerOnTick;
                this.timer.Dispose();
            }

            base.Dispose(disposing);
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this));
            base.Initialize();

            this.dte = (DTE)GetGlobalService(typeof(DTE));
            this.InitializeAddIn();
        }

        /// <summary>
        /// Disconnect from <see cref="DocumentEvents"/>.
        /// </summary>
        private void DisconnectFromVsEvents()
        {
            if (this.documentEvents != null)
            {
                this.documentEvents.DocumentSaved -= this.OnDocumentSaved;
                this.documentEvents.DocumentClosing -= this.OnDocumentClosing;
                this.documentEvents = null;
            }

            if (this.buildEvents != null)
            {
                this.buildEvents.OnBuildBegin -= this.OnBuildBegin;
                this.buildEvents.OnBuildDone -= this.OnBuildDone;
                this.buildEvents = null;
            }

            if (this.solutionEvents != null)
            {
                this.solutionEvents.Opened -= this.OnOpenedSolution;
                this.solutionEvents.BeforeClosing -= this.OnBeforeClosingSolution;
                this.solutionEvents = null;
            }
        }

        /// <summary>
        /// Subscribes to the DocumentSaved event.
        /// </summary>
        private void InitializeAddIn()
        {
            if (this.dte.Commands.Cast<Command>().All(x => x.Name != ReSharperSilentCleanupCodeCommandName))
            {
                return;
            }

            var events2 = this.dte.Events;

            this.DisconnectFromVsEvents();

            this.documentEvents = events2.DocumentEvents[null];
            this.documentEvents.DocumentSaved += this.OnDocumentSaved;
            this.documentEvents.DocumentClosing += this.OnDocumentClosing;

            this.buildEvents = events2.BuildEvents;
            this.buildEvents.OnBuildBegin += this.OnBuildBegin;
            this.buildEvents.OnBuildDone += this.OnBuildDone;

            this.solutionEvents = events2.SolutionEvents;
            this.solutionEvents.Opened += this.OnOpenedSolution;
            this.solutionEvents.BeforeClosing += this.OnBeforeClosingSolution;

            this.timer.Start();
        }

        /// <summary>
        /// The on before closing solution.
        /// </summary>
        private void OnBeforeClosingSolution()
        {
            this.solutionIsActive = false;
        }

        /// <summary>
        /// Called when a build has begun.
        /// </summary>
        /// <param name="scope">
        /// The build scope.
        /// </param>
        /// <param name="action">
        /// The build action.
        /// </param>
        private void OnBuildBegin(vsBuildScope scope, vsBuildAction action)
        {
            switch (action)
            {
                case vsBuildAction.vsBuildActionBuild:
                case vsBuildAction.vsBuildActionRebuildAll:
                case vsBuildAction.vsBuildActionDeploy:
                    ++this.buildingSolution;
                    break;
            }
        }

        /// <summary>
        /// Called when a build is done.
        /// </summary>
        /// <param name="scope">
        /// The build scope.
        /// </param>
        /// <param name="action">
        /// The build action.
        /// </param>
        private void OnBuildDone(vsBuildScope scope, vsBuildAction action)
        {
            switch (action)
            {
                case vsBuildAction.vsBuildActionBuild:
                case vsBuildAction.vsBuildActionRebuildAll:
                case vsBuildAction.vsBuildActionDeploy:
                    --this.buildingSolution;
                    break;
            }
        }

        /// <summary>
        /// Called when a document is being closed.
        /// </summary>
        /// <param name="document">
        /// The document that is being closed.
        /// </param>
        private void OnDocumentClosing(Document document)
        {
            var extension = Path.GetExtension(document.FullName);
            if (!AllowedFileExtensions.Contains(extension))
            {
                return;
            }

            // Do not reformat this document, because at this point, there is no longer a valid window to activate.
            this.documentsToReformatDictionary.Remove(document);
        }

        /// <summary>
        /// Called when a document is saved.
        /// </summary>
        /// <param name="document">
        /// The document that was saved.
        /// </param>
        private void OnDocumentSaved(Document document)
        {
            if (this.isReformatting || this.buildingSolution > 0)
            {
                return;
            }

            var extension = Path.GetExtension(document.FullName);
            if (AllowedFileExtensions.Contains(extension))
            {
                // Mark it for reformatting in our internal data structure
                this.documentsToReformatDictionary[document] = DateTime.Now;
            }
        }

        /// <summary>
        /// Called when a solution is opened.
        /// </summary>
        private void OnOpenedSolution()
        {
            this.solutionIsActive = true;
        }

        /// <summary>
        /// Reformats a given list of documents, ensuring that the background timer is stopped during the process.
        /// </summary>
        /// <param name="documentsToReformat">
        /// The documents to reformat.
        /// </param>
        /// <param name="saveDocumentsAfterwards">
        /// <c>true</c> if the changed documents should be saved afterwards.
        /// </param>
        private void ReformatDocuments(IEnumerable<Document> documentsToReformat, bool saveDocumentsAfterwards = false)
        {
            this.isReformatting = true;
            this.timer.Stop();
            this.lastReformat = DateTime.Now;

            var originallyActiveWindow = this.dte.ActiveWindow;

            try
            {
                var originallyActiveDocument = originallyActiveWindow != null
                    ? originallyActiveWindow.Document
                    : null;
                var activeDocumentCollection = originallyActiveDocument != null
                    ? Enumerable.Repeat(originallyActiveDocument, 1)
                    : Enumerable.Empty<Document>();

// ReSharper disable PossibleMultipleEnumeration
                var recentlySavedDocs =
                    documentsToReformat // .Where(x => x.ActiveWindow != null)
                        .Except(activeDocumentCollection)
                        .Concat(originallyActiveDocument != null && this.documentsToReformatDictionary.ContainsKey(originallyActiveDocument)
                            ? activeDocumentCollection
                            : Enumerable.Empty<Document>())



// ReSharper restore PossibleMultipleEnumeration
                        .ToArray();

                foreach (var document in recentlySavedDocs)
                {
                    // Active the document which was just saved
                    document.Activate();

                    try
                    {
                        // so that we can run the ReSharper command on it.
                        this.dte.ExecuteCommand(ReSharperSilentCleanupCodeCommandName);
                    }
                    catch (Exception)
                    {
                    }
                    finally
                    {
                        this.documentsToReformatDictionary.Remove(document);
                    }
                }

                if (saveDocumentsAfterwards)
                {
                    foreach (var document in recentlySavedDocs.Where(document => !document.Saved))
                    {
                        document.Save();
                    }
                }

                if (originallyActiveWindow != null)
                {
                    // Reactivate the original window.
                    originallyActiveWindow.Activate();
                }
            }
            finally
            {
                foreach (var document in documentsToReformat)
                {
                    this.documentsToReformatDictionary.Remove(document);
                }

                this.timer.Start();
                this.isReformatting = false;
            }
        }

        private static bool IsVisualStudioForegroundWindow()
        {
            uint foregroundProcessId;
            NativeMethods.GetWindowThreadProcessId(NativeMethods.GetForegroundWindow(), out foregroundProcessId);
            var visualStudioProcessId = System.Diagnostics.Process.GetCurrentProcess().Id;
            return visualStudioProcessId == foregroundProcessId;
        }

        /// <summary>
        /// Called periodically.
        /// </summary>
        /// <param name="sender">
        /// The timer.
        /// </param>
        /// <param name="eventArgs">
        /// Nothing (<see cref="EventArgs.Empty"/>).
        /// </param>
        private void TimerOnTick(object sender, EventArgs eventArgs)
        {
            try
            {
                if (this.buildingSolution > 0)
                {
                    this.documentsToReformatDictionary.Clear();

                    return;
                }

                // Sanity check
                if (this.dte.Application.Mode == vsIDEMode.vsIDEModeDebug ||
                    this.isReformatting ||
                    !this.solutionIsActive ||
                    !this.documentsToReformatDictionary.Any() ||
                    !IsVisualStudioForegroundWindow())
                {
                    return;
                }

                // Remove all unsaved documents from the dictionary
                foreach (var document in this.documentsToReformatDictionary.Where(x => !x.Key.Saved).Select(x => x.Key).ToArray())
                {
                    this.documentsToReformatDictionary.Remove(document);
                }

                var now = DateTime.Now;
                if ((now - this.lastReformat).TotalSeconds < 5)
                {
                    // Ignore any documents that have been saved if a reformat has happened within the last 5 seconds.
                    this.documentsToReformatDictionary.Clear();
                    return;
                }

                var anyDocumentSavedSinceLastCheck = this.documentsToReformatDictionary.Any(kvp => (now - kvp.Value).TotalMilliseconds < this.timer.Interval);
                if (!this.documentsToReformatDictionary.Any() || anyDocumentSavedSinceLastCheck)
                {
                    return;
                }

                this.ReformatDocuments(
                    this.documentsToReformatDictionary
                        .OrderBy(kvp => kvp.Value)
                        .Select(kvp => kvp.Key),
                    true);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(string.Format("{0}\n\n{1}", e.Message, e.StackTrace), "ReShaper.AutoFormatOnSave", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}