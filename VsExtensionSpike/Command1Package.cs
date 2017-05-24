//------------------------------------------------------------------------------
// <copyright file="Command1Package.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Utilities;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace VsExtensionSpike
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
    [ProvideToolWindow(typeof(SpikeWindow))]
    [Guid(Command1Package.PackageGuidString)]
    [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class Command1Package : Package, IOleCommandTarget
    {
        /// <summary>
        /// Command1Package GUID string.
        /// </summary>
        public const string PackageGuidString = "979735ed-4c85-48c9-b0e9-5b274d1be907";

        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// </summary>
        public Command1Package()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Command1.Initialize(this);
            SpikeWindowCommand.Initialize(this);
            base.Initialize();
        }

        #endregion

        public int QueryStatus(ref Guid guidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            Debug.Assert(cCmds == 1, "Multiple commands");
            Debug.Assert(prgCmds != null, "NULL argument");

            if ((prgCmds == null))
            {
                return VSConstants.E_INVALIDARG;
            }

            for (int i = 0; i < cCmds; i++)
            {
                OLECMDF cmdf = OLECMDF.OLECMDF_SUPPORTED;
                var command = prgCmds[i];
                switch (command.cmdID)
                {
                    case Command1.CommandId:
                        if (Command1.Instance != null && Command1.Instance.IsAvailable())
                        {
                            cmdf |= OLECMDF.OLECMDF_ENABLED;
                        }
                        else
                        {
                            cmdf |= OLECMDF.OLECMDF_INVISIBLE;
                        }
                        prgCmds[i].cmdf = (uint)cmdf;
                        break;
                    case SpikeWindowCommand.CommandId:
                        //cmdf |= OLECMDF.OLECMDF_INVISIBLE;
                        //prgCmds[i].cmdf = (uint)cmdf;
                        break;

                    default:
                        break;
                }
            }

            return VSConstants.S_OK;
        }
    }
}
