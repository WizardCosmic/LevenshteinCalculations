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
using LevenshteinCalculations;

namespace LevenshteinCalculations
{
    /// <summary>
    /// Interaction logic for CalcWindow.xaml
    /// </summary>
    /// 
    
    public partial class CalcWindow : Window
    {
        private WordPair pair;
        public CalcWindow(WordPair wp)
        {
            InitializeComponent();
            pair  = wp;
        }
    }

    
}
