//------------------------------------------------------------------------------
// <copyright file="ScriptorPackage.cs" company="LabSolvay">
//     Copyright (c) LabSolvay.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace SSMS2017_AddIn
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string)]
    [Guid(ScriptorPackage.PackageGuidString)]
    [ProvideAutoLoad(UIContextGuids80.NoSolution)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class ScriptorPackage : Package
    {
        /// <summary>
        /// ScriptorPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "514c27e3-bc7d-4f8b-a3b7-0b6c89e3e82c";
        public static DTE2 applicationObject;
        private static DteInitializer dteInitializer;
    
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Scriptor.Initialize(this);
            base.Initialize();
            InitializeDTE();
            AddSkipLoading();
        }

        private void AddSkipLoading()
        {
            var timer = new System.Timers.Timer(2000);
            timer.Elapsed += (sender, args) =>
            {
                timer.Stop();
                var myPackage = UserRegistryRoot.CreateSubKey(@"Packages\{" + PackageGuidString + "}");
                myPackage?.SetValue("SkipLoading", 1);
            };
            timer.Start();
        }

        internal class DteInitializer : IVsShellPropertyEvents
        {
            private readonly IVsShell shellService;
            private uint cookie;
            private readonly Action callback;

            internal DteInitializer(IVsShell shellService, Action callback)
            {
                this.shellService = shellService;
                this.callback = callback;

                // Set an event handler to detect when the IDE is fully initialized
                var hr = this.shellService.AdviseShellPropertyChanges(this, out this.cookie);

                ErrorHandler.ThrowOnFailure(hr);
            }

            int IVsShellPropertyEvents.OnShellPropertyChange(int propid, object var)
            {
                if (propid == (int)__VSSPROPID.VSSPROPID_Zombie)
                {
                    var isZombie = (bool)var;

                    if (!isZombie)
                    {
                        // Release the event handler to detect when the IDE is fully initialized
                        var hr = this.shellService.UnadviseShellPropertyChanges(this.cookie);

                        ErrorHandler.ThrowOnFailure(hr);

                        this.cookie = 0;

                        this.callback();
                    }
                }
                return VSConstants.S_OK;
            }
        }
        /// <summary>
        /// to get the reference of IDE (VS or SSMS)
        /// </summary>
        private void InitializeDTE()
        {
            applicationObject = this.GetService(typeof(SDTE)) as EnvDTE80.DTE2;

            if (applicationObject == null) // The IDE is not yet fully initialized
            {
                var shellService = this.GetService(typeof(SVsShell)) as IVsShell;
                dteInitializer = new DteInitializer(shellService, this.InitializeDTE);
            }
            else
            {
                dteInitializer = null;
            }
        }
        #endregion
    }
}
