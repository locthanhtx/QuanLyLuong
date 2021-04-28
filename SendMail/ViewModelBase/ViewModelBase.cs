using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendMail
{
    public abstract class ViewModelBase : BindableBase
    {
        public string Name { get; set; }

        #region IsBusy

        private bool _isBusy;
        /// <summary>
        /// Gets or sets IsBusy
        /// </summary>
        public bool IsBusy
        {
            get
            {
                return _isBusy;
            }
            set
            {
                if (_isBusy != value)
                {
                    _isBusy = value;
                    OnPropertyChanged(() => IsBusy);
                }
            }
        }

        #endregion

    }
}
