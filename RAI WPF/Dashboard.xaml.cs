using System;
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
using System.Windows.Shapes;

namespace RAI_WPF
{
    /// <summary>
    /// Interaction logic for Dashboard.xaml
    /// </summary>
    public partial class Dashboard : Window
    {
        public Dashboard()
        {
            InitializeComponent();
            //this.Browser.Navigate("https://app.powerbi.com/view?r=eyJrIjoiZmZmYzliZWUtMzA2OS00ZTFiLWJkOGUtMmJmNWExYzcyNjA1IiwidCI6IjcyOWQxY2U4LTUxN2EtNGJjYi1iY2JlLTgzYThmNzJkYzMzMSIsImMiOjF9/");
            this.Browser.Navigate("https://app.powerbi.com/view?r=eyJrIjoiZTYxNTc0YTEtOGFhNi00NmIzLWE2MmQtODQ3NDM4Yzc1NzU4IiwidCI6IjcyOWQxY2U4LTUxN2EtNGJjYi1iY2JlLTgzYThmNzJkYzMzMSIsImMiOjF9&pageName=ReportSection/");
        }
    }
}
