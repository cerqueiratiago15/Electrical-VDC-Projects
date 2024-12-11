using System;
using System.Windows.Input;

namespace ElectricalSolutions
{
    public class RelayCommand : ICommand
    {
        public event EventHandler CanExecuteChanged = (sender, e) => { };

        private Action mAction;

        public RelayCommand(Action action)
        {
            this.mAction = action;
        }
        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            this.mAction();
        }
    }
}
