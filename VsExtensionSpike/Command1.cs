﻿//------------------------------------------------------------------------------
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
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="pckg">Owner package, not null.</param>
        private Command1(Package pckg)
        {
            package = pckg ?? throw new ArgumentNullException(nameof(pckg));

            if (this.ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                menuCommand = menuItem;
                Microsoft.VisualStudio.Text.Editor.IWpfTextView textView = GetTextView();
                textView.Caret.PositionChanged += Caret_PositionChanged;
                commandService.AddCommand(menuItem);

                menuCommandID = new CommandID(CommandSet, SolutionCommandId);
                menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        private void Caret_PositionChanged(object sender, Microsoft.VisualStudio.Text.Editor.CaretPositionChangedEventArgs e)
        {
            Microsoft.VisualStudio.Text.Editor.IWpfTextView textView = e.TextView as Microsoft.VisualStudio.Text.Editor.IWpfTextView;

            try
            {
                Microsoft.VisualStudio.Text.SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;
                var document = Microsoft.CodeAnalysis.Text.Extensions.GetOpenDocumentInCurrentContextWithChanges(caretPosition.Snapshot);
                var node = document.GetSyntaxRootAsync().Result.FindToken(caretPosition).Parent;

                if (node is MethodDeclarationSyntax)
                {
                    menuCommand.Visible = true;
                    menuCommand.Enabled = true;
                }
                else
                {
                    menuCommand.Visible = false;
                    menuCommand.Enabled = false;
                }
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.ToString());
                menuCommand.Visible = false;
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
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

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
            var oleMenuCommand = sender as OleMenuCommand;
            //oleMenuCommand.BeforeQueryStatus += OleMenuCommand_BeforeQueryStatus;

            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "Command1";

            // Show a message box to prove we were here
            //VsShellUtilities.ShowMessageBox(
            //    this.ServiceProvider,
            //    message,
            //    title,
            //    OLEMSGICON.OLEMSGICON_INFO,
            //    OLEMSGBUTTON.OLEMSGBUTTON_OK,
            //    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            Microsoft.VisualStudio.Text.Editor.IWpfTextView textView = GetTextView();
            Microsoft.VisualStudio.Text.SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;
            var document = Microsoft.CodeAnalysis.Text.Extensions.GetOpenDocumentInCurrentContextWithChanges(caretPosition.Snapshot);
            var node = document.GetSyntaxRootAsync().Result.FindToken(caretPosition).Parent;



            //MethodDeclarationSyntax testMethod =
            //    document.GetSyntaxRootAsync().Result.FindToken(caretPosition).Parent.AncestorsAndSelf()
            //    .OfType<MethodDeclarationSyntax>().FirstOrDefault();
        }

        private void OleMenuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {
            Microsoft.VisualStudio.Text.Editor.IWpfTextView textView = GetTextView();
            Microsoft.VisualStudio.Text.SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;
            var document = Microsoft.CodeAnalysis.Text.Extensions.GetOpenDocumentInCurrentContextWithChanges(caretPosition.Snapshot);
            var node = document.GetSyntaxRootAsync().Result.FindToken(caretPosition).Parent;

            if (node is MethodDeclarationSyntax)
            {
                menuCommand.Visible = true;
            }
            else
            {
                menuCommand.Visible = false;
            }
            throw new NotImplementedException();
        }

        private Microsoft.VisualStudio.Text.Editor.IWpfTextView GetTextView()
        {
            IVsTextManager textManager = (IVsTextManager)ServiceProvider.GetService(typeof(SVsTextManager));
            textManager.GetActiveView(1, null, out IVsTextView textView);
            return GetEditorAdaptersFactoryService().GetWpfTextView(textView);
        }
        private Microsoft.VisualStudio.Editor.IVsEditorAdaptersFactoryService GetEditorAdaptersFactoryService()
        {
            Microsoft.VisualStudio.ComponentModelHost.IComponentModel componentModel =
                ServiceProvider.GetService(typeof(Microsoft.VisualStudio.ComponentModelHost.SComponentModel)) as Microsoft.VisualStudio.ComponentModelHost.IComponentModel;
            return componentModel.GetService<Microsoft.VisualStudio.Editor.IVsEditorAdaptersFactoryService>();
        }
    }
}
