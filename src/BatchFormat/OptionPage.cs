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
        private string _excludePath;

        [Description(@"Exclude files when path contains
eg: AssemblyInfo.cs|.designer.cs|Reference.cs|\metadata\
split by '|' £¬IgnoreCase

If Empty will use the default values.")]
        public string ExcludePath
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_excludePath))
                {
                    return string.Join("|", ConfigHelper.ExcludePathList);
                }
                return _excludePath;
            }
            set
            {
                _excludePath = value;
            }
        }

    }
}
