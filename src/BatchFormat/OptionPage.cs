using System;
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

        [Description(@"Default:AssemblyInfo.cs|.designer.cs|Reference.cs|\MetaData\
mean:Exclude files when path contains these,split by '|',IgnoreCase

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

        [Description(@"Default:true Mean£º
if <Current>.cs is gen by *.tt,then Ignore")]
        public bool IgnoreT4Child { get; set; } = true;

    }
}
