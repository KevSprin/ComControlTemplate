using RegFreeActiveXControls;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ComControlTemplate
{
    [ComVisible(true)]
    [Guid("00213331-FB5D-4936-8BDC-B49850FA4E21"), ClassInterface(ClassInterfaceType.None)]
    public partial class ComControlTemplate: UserControl, IExposedProperties
    {
        public ComControlTemplate()
        {
            InitializeComponent();
        }

        public string CustomText { get; set; } = string.Empty;

        [EditorBrowsable(EditorBrowsableState.Always)]
        [ComRegisterFunction]
        private static void Register(Type t)
        {
            ComRegistration.RegisterControl(t);
        }

        [EditorBrowsable(EditorBrowsableState.Always)]
        [ComUnregisterFunction]
        private static void Unregister(Type t)
        {
            ComRegistration.UnregisterControl(t);
        }
    }
}
