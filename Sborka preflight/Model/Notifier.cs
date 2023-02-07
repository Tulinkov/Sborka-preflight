using System;
using System.ComponentModel;

namespace SborkaPreflight.Model
{
    class Notifier : INotifyPropertyChanged
    {
        #region INotyfyPropertyChanged members

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion
    }
}
