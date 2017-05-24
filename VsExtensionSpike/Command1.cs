//------------------------------------------------------------------------------
// <copyright file="Command1.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using EnvDTE;
using EnvDTE80;
using System.Linq;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;

namespace VsExtensionSpike
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;
        public const int SolutionCommandId = 0x0020;
        public readonly OleMenuCommandService commandService;
        public readonly MenuCommand menuCommand;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("6fcab3de-9e67-4037-838f-0bd763553dd3");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;


        /// <summary>
        /// 
        /// </summary>
        private MethodDeclarationSyntax MethodDeclarationSyntax { get; set; } = null;


        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="pckg">Owner package, not null.</param>
        private Command1(Package pckg)
        {
            package = pckg ?? throw new ArgumentNullException(nameof(pckg));

            if (this.ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                this.commandService = commandService;
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
                menuCommand = menuItem;
                commandService.AddCommand(menuItem);
            }

            /*
             * Available commands: open command window> Tools.$AvailableCommand$
             * var dte = (DTE2)ServiceProvider.GetService(typeof(DTE)); an example of how to get DTE
             file1
             file2
              dte.ExecuteCommand("Tools.$AvailableCommand$", $"\"{file1}\" \"{file2}\""; example of how to execute existing commands in VS
             */
        }

        /// <summary>
        /// 
        /// </summary>
        private void DisableCommand()
        {
            if (menuCommand.Enabled || menuCommand.Visible)
            {
                menuCommand.Visible = false;
                menuCommand.Enabled = false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void EnableCommand()
        {
            if (!menuCommand.Enabled || !menuCommand.Visible)
            {
                menuCommand.Enabled = true;
                menuCommand.Visible = true;
            }
        }


        /// <summary>
        /// Gets the instance of the command.
        /// Enables a VSPackage to gain access to IntelliMouse functionality such as using the mouse wheel and handling scroll and pan bitmaps when the mouse wheel is clicked.
        /// </summary>
        public static Command1 Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        public IServiceProvider ServiceProvider { get => package; }

        [Import]
        internal IClassifierAggregatorService classifierAggregatorService = null;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new Command1(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
            dte.ExecuteCommand("View.SpikeWindow");
            SpikeWindowCommand.Instance.Test = this.MethodDeclarationSyntax;
            Console.Write("");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private Microsoft.VisualStudio.Text.Editor.IWpfTextView GetTextView()
        {
            IVsTextManager textManager = (IVsTextManager)ServiceProvider.GetService(typeof(SVsTextManager));
            textManager.GetActiveView(1, null, out IVsTextView textView);
            return GetEditorAdaptersFactoryService().GetWpfTextView(textView);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private Microsoft.VisualStudio.Editor.IVsEditorAdaptersFactoryService GetEditorAdaptersFactoryService()
        {
            Microsoft.VisualStudio.ComponentModelHost.IComponentModel componentModel =
                ServiceProvider.GetService(typeof(Microsoft.VisualStudio.ComponentModelHost.SComponentModel)) as Microsoft.VisualStudio.ComponentModelHost.IComponentModel;
            return componentModel.GetService<Microsoft.VisualStudio.Editor.IVsEditorAdaptersFactoryService>();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool IsAvailable()
        {
            try
            {
                Microsoft.VisualStudio.Text.Editor.IWpfTextView textView = GetTextView();
                Microsoft.VisualStudio.Text.SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;

                var contentType = caretPosition.Snapshot.ContentType;
                if (String.Equals(contentType.TypeName, @"CSharp", StringComparison.Ordinal))
                {
                    var document = Microsoft.CodeAnalysis.Text.Extensions.GetOpenDocumentInCurrentContextWithChanges(caretPosition.Snapshot);
                    var node = document.GetSyntaxRootAsync().Result.FindToken(caretPosition).Parent;
                    if (node is MethodDeclarationSyntax selected)
                    {
                        MethodDeclarationSyntax = selected;
                        EnableCommand();
                        return true;
                    }
                }
                DisableCommand();
            }
            catch (Exception)
            {
                DisableCommand();
            }

            MethodDeclarationSyntax = null;
            return false;
        }
    }
}
