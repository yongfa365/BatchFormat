/***************************************************************************

Copyright (c) Microsoft Corporation. All rights reserved.
This code is licensed under the Visual Studio SDK license terms.
THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.

***************************************************************************/

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace YongFa365.BatchFormat
{
    [Guid(GuidList.GuidPageGeneralString)]
    [ComVisible(true)]
    public class OptionsPageGeneral : DialogPage
    {
        [Description("Set EndsWith Array. \r\nIf Empty will use the default values.")]
        public string[] EndsWith { get; set; }

        protected override void OnActivate(CancelEventArgs e)
        {
            EndsWith = ConfigHelper.Load().ToArray();
            base.OnActivate(e);
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            ConfigHelper.Save(EndsWith);
            base.OnApply(e);
        }
    }
}
