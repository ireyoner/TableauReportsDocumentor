﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TableauReportsDocumentor.Data;

namespace TableauReportsDocumentor.Modules.EditModule
{
    /// <summary>
    /// Interaction logic for Editor.xaml
    /// </summary>
    public partial class Editor : UserControl
    {
        public ReportContent Report
        {
            set
            {
                RTV.Report = value;
                RXV.Report = value;
                ROV.Report = value;
            }
        }

        public Label statusLabel
        {
            set
            {
                RXV.statusLabel = value;
                RTV.statusLabel = value;
            }
        }

        public Editor()
        {
            InitializeComponent();
        }
    }
}
