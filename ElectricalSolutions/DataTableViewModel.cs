using System.Data;
using System.Windows.Input;

namespace ElectricalSolutions
{
    // ViewModel
    public class DataTableViewModel : Notifier
    {
        public DataTable Data { get; set; }

        public ICommand ImportCommand { get; private set; }
        public ICommand ExportCommand { get; private set; }

        public DataTableViewModel()
        {
            ImportCommand = new RelayCommand(ImportData);
            ExportCommand = new RelayCommand(ExportData);
            
        }

        private void ImportData()
        {
            // Implement import logic here
        }

        private void ExportData()
        {
            // Implement export logic here
        }
    }

}





