/***************************************************************************

Copyright (c) Microsoft Corporation. All rights reserved.
This code is licensed under the Visual Studio SDK license terms.
THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.

***************************************************************************/

using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace YongFa365.BatchFormat
{
    [Guid(GuidList.GuidPageGeneralString)]
    [ComVisible(true)]
    public class OptionPage : DialogPage
    {
        private string excludePath;

        [Description("Exclude files when path contains eg: EndsWith AssemblyInfo.cs|.designer.cs|Reference.cs|/metadata/ \r\nsplit by \"|\" \r\n\r\nIf Empty will use the default values.")]
        public string ExcludePath
        {
            get
            {
                if (string.IsNullOrWhiteSpace(excludePath))
                {
                    return string.Join("|", ConfigHelper.ExcludePathList);
                }
                return excludePath;
            }
            set
            {
                excludePath = value;
            }
        }

    }
}
