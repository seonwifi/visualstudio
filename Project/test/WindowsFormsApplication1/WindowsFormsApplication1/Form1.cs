using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Unity\Editor\Unity.exe", @"-quit -batchmode -projectPath D:\UnityProject\BladeAxe -executeMethod SmBatchBuild.BuildApk");
        }
    }
}
