using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using Forms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


namespace SpellCheckWordAddIn
{
    /// <summary>
    /// Interaction logic for SpellCheckUserControl.xaml
    /// </summary>
    public partial class SpellCheckUserControl : UserControl
    {
        private ObservableCollection<SpellError> spellErrors;

        public SpellCheckUserControl()
        {
            InitializeComponent();

            spellErrors = new ObservableCollection<SpellError>();
            spellErrorsItemsControl.ItemsSource = spellErrors;

            this.Loaded += SpellCheckUserControl_Loaded;
        }

        private void SpellCheckUserControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                Load();
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void refreshButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Load();
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                for (int i = 0; i < spellErrors.Count; i++)
                {
                    spellErrors[i].Range.Text = spellErrors[i].Text;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void gotoButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Button button = (Button)sender;
                if (button.DataContext is SpellError spellError)
                {
                    spellError.Range.Select();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public void Load()
        {
            try
            {
                spellErrors.Clear();

                Word.Application application = Globals.ThisAddIn.Application;
                Word.Document activeDocument = application.ActiveDocument;
                for (int i = 1; i <= activeDocument.Content.SpellingErrors.Count; i++)
                {
                    SpellError spellError = new SpellError()
                    {
                        Text = activeDocument.Content.SpellingErrors[i].Text,
                        Range = activeDocument.Content.SpellingErrors[i],
                    };

                    spellErrors.Add(spellError);
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

       
    }


    public class SpellError
    {
        public string Text { get; set; }

        public Word.Range Range { get; set; }
    }
}
