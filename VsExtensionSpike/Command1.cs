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
        private ITextCaret caret;

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
                this.commandService = commandService;
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
                menuCommand = menuItem;
                //menuCommand.BeforeQueryStatus += MenuCommand_BeforeQueryStatus;
                commandService.AddCommand(menuItem);
            }
        }

        private void MenuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void Caret_PositionChanged(object sender, Microsoft.VisualStudio.Text.Editor.CaretPositionChangedEventArgs e)
        {
            return;
            Microsoft.VisualStudio.Text.Editor.IWpfTextView textView = e.TextView as Microsoft.VisualStudio.Text.Editor.IWpfTextView;
            Microsoft.VisualStudio.Text.SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;
            var document = Microsoft.CodeAnalysis.Text.Extensions.GetOpenDocumentInCurrentContextWithChanges(caretPosition.Snapshot);
            var contentType = caretPosition.Snapshot.ContentType;
            if (!String.Equals(contentType.TypeName, @"CSharp", StringComparison.Ordinal))
            {
                DisableCommand();
                return;
            }
            try
            {
                var node = document.GetSyntaxRootAsync().Result.FindToken(caretPosition).Parent;
                if (node is MethodDeclarationSyntax)
                {
                    EnableCommand();
                }
                else
                {
                    DisableCommand();
                }
            }
            catch (Exception)
            {
                DisableCommand();
            }
        }

        private void DisableCommand()
        {
            if (menuCommand.Enabled || menuCommand.Visible)
            {
                menuCommand.Visible = false;
                menuCommand.Enabled = false;
            }
        }

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
        private IServiceProvider ServiceProvider { get => this.package; }

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
        /// 
        /// 
        /// sender as OleMenuCommand
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
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


        public bool IsAvailable()
        {
            try
            {
                Microsoft.VisualStudio.Text.Editor.IWpfTextView textView = GetTextView();
                Microsoft.VisualStudio.Text.SnapshotPoint caretPosition = textView.Caret.Position.BufferPosition;

                //var currentCaret = textView.Caret;
                //if (caret == null)
                //{
                //    caret = currentCaret;
                //    currentCaret.PositionChanged += Caret_PositionChanged;
                //}
                //else
                //{
                //    if(currentCaret != caret)
                //    {
                //        currentCaret.PositionChanged += Caret_PositionChanged;
                //    }
                //}

                var contentType = caretPosition.Snapshot.ContentType;
                if (String.Equals(contentType.TypeName, @"CSharp", StringComparison.Ordinal))
                {
                    var document = Microsoft.CodeAnalysis.Text.Extensions.GetOpenDocumentInCurrentContextWithChanges(caretPosition.Snapshot);
                    var node = document.GetSyntaxRootAsync().Result.FindToken(caretPosition).Parent;
                    if (node is MethodDeclarationSyntax)
                    {
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
            return false;
        }
    }
}
