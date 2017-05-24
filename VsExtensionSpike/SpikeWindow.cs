//------------------------------------------------------------------------------
// <copyright file="SpikeWindow.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace VsExtensionSpike
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.CodeAnalysis.CSharp.Syntax;

    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("617abb78-6d3a-490b-92ea-c12b6ee7e681")]
    public class SpikeWindow : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SpikeWindow"/> class.
        /// </summary>
        public SpikeWindow() : base(null)
        {
            this.Caption = "SpikeWindow";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new SpikeWindowControl();
        }

        public void UpdateMe(MethodDeclarationSyntax mds)
        {
            var spikeWindowControl = (SpikeWindowControl)this.Content;

            if (mds != null)
            {
                spikeWindowControl.UpdateMe(mds.Identifier.ValueText);
            }
        }
    }
}
