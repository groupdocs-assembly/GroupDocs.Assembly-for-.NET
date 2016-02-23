// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace GroupDocsAssemblyVisualStudioPlugin.Core
{
    public class GroupDocsComponents
    {
        public static Dictionary<String, GroupDocsComponent> list = new Dictionary<string, GroupDocsComponent>();
        public GroupDocsComponents()
        {
            list.Clear();

            GroupDocsComponent groupdocsassembly = new GroupDocsComponent();
            groupdocsassembly.set_downloadUrl("");
            groupdocsassembly.set_downloadFileName("groupdocs.assembly.zip");
            groupdocsassembly.set_name(Constants.GROUPDOCS_COMPONENT);
            groupdocsassembly.set_remoteExamplesRepository("https://github.com/groupdocsassembly/GroupDocs_Assembly_NET.git");
            list.Add(Constants.GROUPDOCS_COMPONENT, groupdocsassembly);
        }
    }
}
