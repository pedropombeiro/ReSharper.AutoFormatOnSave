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

    using EnvDTE80;

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
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
// This attribute is used to register the informations needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidReSharper_AutoFormatOnSavePkgString)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
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
        private static readonly string[] AllowedFileExtensions = new[] { ".cs", ".xaml", ".vb" };

        #endregion

        #region Fields

        /// <summary>
        /// The dictionary of documents which are currently being reformatted mapped to the timestamp when they were last reformatted.
        /// </summary>
        private readonly Dictionary<Document, DateTime> recentlySavedDocsDictionary = new Dictionary<Document, DateTime>();

        /// <summary>
        /// A timer which runs background checks (mainly to check whether all saves have completed).
        /// </summary>
        private readonly Timer timer = new System.Windows.Forms.Timer { Interval = 500 };

        /// <summary>
        /// The document events object.
        /// </summary>
        private DocumentEvents documentEvents;

        /// <summary>
        /// The DTE2 global service.
        /// </summary>
        private DTE2 dte;

        /// <summary>
        /// <c>true</c> if currently reformatting recently saved files.
        /// </summary>
        private bool isReformatting;

        /// <summary>
        /// The time of the last reformat.
        /// </summary>
        private DateTime lastReformat;

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
                this.DisconnectFromDocumentEvents();

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

            this.dte = (DTE2)GetGlobalService(typeof(DTE));
            this.InitializeAddIn();
        }

        /// <summary>
        /// Disconnect from <see cref="DocumentEvents"/>.
        /// </summary>
        private void DisconnectFromDocumentEvents()
        {
            if (this.documentEvents == null)
            {
                return;
            }

            this.documentEvents.DocumentSaved -= this.OnDocumentSaved;
            this.documentEvents = null;
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

            var events2 = (EnvDTE80.Events2)this.dte.Events;

            this.DisconnectFromDocumentEvents();

            this.documentEvents = events2.DocumentEvents[null];
            this.documentEvents.DocumentSaved += this.OnDocumentSaved;
            this.timer.Start();
        }

        /// <summary>
        /// Called when a document is saved.
        /// </summary>
        /// <param name="document">
        /// The document that was saved.
        /// </param>
        private void OnDocumentSaved(Document document)
        {
            var extension = Path.GetExtension(document.FullName);
            if (AllowedFileExtensions.Contains(extension))
            {
                // Mark it as saved in our internal data structure
                this.recentlySavedDocsDictionary[document] = DateTime.Now;
            }
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
            // Sanity check
            if (this.isReformatting || !this.recentlySavedDocsDictionary.Any())
            {
                return;
            }

            if ((DateTime.Now - this.lastReformat).TotalSeconds < 5)
            {
                // Ignore any documents that have been saved if a reformat has happened within the last 5 seconds.
                this.recentlySavedDocsDictionary.Clear();
                return;
            }

            var anyDocumentSavedSinceLastCheck = this.recentlySavedDocsDictionary.Any(kvp => (DateTime.Now - kvp.Value).TotalMilliseconds < this.timer.Interval);
            if (anyDocumentSavedSinceLastCheck)
            {
                return;
            }

            this.isReformatting = true;
            this.timer.Stop();

            var originallyActiveWindow = this.dte.ActiveWindow;

            try
            {
                var recentlySavedDocs =
                    this.recentlySavedDocsDictionary
                        .OrderBy(kvp => kvp.Value)
                        .Select(kvp => kvp.Key)
                        .ToArray();

                foreach (var document in recentlySavedDocs)
                {
                    // Active the document which was just saved
                    document.Activate();

                    // so that we can run the ReSharper command on it.
                    this.dte.ExecuteCommand(ReSharperSilentCleanupCodeCommandName);
                }

                foreach (var document in recentlySavedDocs)
                {
                    // and save it again, if changed.
                    if (!document.Saved)
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
                this.recentlySavedDocsDictionary.Clear();
                this.lastReformat = DateTime.Now;
                this.timer.Start();
                this.isReformatting = false;
            }
        }

        #endregion
    }
}