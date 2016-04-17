using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace QueryXLerator
{
    public class ViewModelBase : INotifyPropertyChanged
    {
        private readonly Dictionary<string, object> values = new Dictionary<string, object>();

        public event PropertyChangedEventHandler PropertyChanged;

        internal T GetValue<T>([CallerMemberName]string member = "")
        {
            object currentObjectValue = null;
            if (values.TryGetValue(member, out currentObjectValue))
            {
                return (T)currentObjectValue;
            }
            return default(T);
        }

        internal void SetValue<T>(T newValue, [CallerMemberName]string member = "")
        {
            T currentValue = GetValue<T>(member);

            bool theValuesAreDifferent = object.Equals(currentValue, newValue) == false;
            if (theValuesAreDifferent == false)
            {
                return;
            }

            // always set the new value
            values[member] = newValue;
            RaisePropChanged(member);
        }

        protected void RaisePropChanged(string membername)
        {
            if (PropertyChanged == null)
            {
                return;
            }
            PropertyChanged(this, new PropertyChangedEventArgs(membername));
        }
    }
}