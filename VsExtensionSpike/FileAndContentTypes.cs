using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition.Primitives;

namespace VsExtensionSpike
{
        internal static class FileAndContentTypeDefinitions
        {
        internal const string _name = "TestAnalyzer";

        [Export]
        [Name(_name)]
        [BaseDefinition("code")]
        internal static ContentTypeDefinition vsContentTypeDefinition;

        [Export]
        [FileExtension(".cs")]
        [ContentType(_name)]
        internal static FileExtensionToContentTypeDefinition csFileExtensionDefinition;

    }
}
