using System.ComponentModel;

namespace QueryXLerator
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void RaisePropChanged(string membername)
        {
            if (PropertyChanged == null)
            {
                return;
            }
            PropertyChanged(this, new PropertyChangedEventArgs(membername));
        }

        //internal void SetValue<T>(T oldValue, T newValue, [CallerMemberName]string member = "") where T : IEquatable<T>
        //{
        //    if (PropertyChanged == null)
        //    {
        //        return;
        //    }
        //    if (oldValue.Equals(newValue))
        //    {
        //        return;
        //    }
        //    PropertyChanged(this, new PropertyChangedEventArgs(member));
        //}
    }
}