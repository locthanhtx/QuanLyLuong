using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace TinhThuong
{
    /// <summary>
    /// A command whose sole purpose is to relay its functionality to other objects by invoking delegates. 
    /// The default return value for the CanExecute method is 'true'.
    /// </summary>
    /// <typeparam name="T">General type</typeparam>
    public class RelayCommand<T> : ICommand
    {
        /// <summary>
        /// Can execute function.
        /// </summary>
        private readonly Predicate<T> canExecute;

        /// <summary>
        /// Execute command action.
        /// </summary>
        private readonly Action<T> execute;

        /// <summary>
        /// Initializes a new instance of the <see cref="RelayCommand&lt;T&gt;"/> class and the command can always be executed.
        /// </summary>
        /// <param name="execute">The execution logic.</param>
        public RelayCommand(Action<T> execute)
            : this(execute, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RelayCommand&lt;T&gt;"/> class.
        /// </summary>
        /// <param name="execute">The execution logic.</param>
        /// <param name="canExecute">The execution status logic.</param>
        public RelayCommand(Action<T> execute, Predicate<T> canExecute)
        {
            if (execute == null)
            {
                throw new ArgumentNullException("execute");
            }

            this.execute = execute;
            this.canExecute = canExecute;
        }

        /// <summary>
        /// Can execute event.
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add
            {
                if (this.canExecute != null)
                {
                    CommandManager.RequerySuggested += value;
                }
            }

            remove
            {
                if (this.canExecute != null)
                {
                    CommandManager.RequerySuggested -= value;
                }
            }
        }

        /// <summary>
        /// Can execute command.
        /// </summary>
        /// <param name="parameter">Attached parameter</param>
        /// <returns>Return true if can execute otherwise return false.</returns>
        public bool CanExecute(object parameter)
        {
            return this.canExecute == null ? true : this.canExecute((T)parameter);
        }

        /// <summary>
        /// Execute command
        /// </summary>
        /// <param name="parameter">Attached parameter</param>
        public void Execute(object parameter)
        {
            this.execute((T)parameter);
        }
    }

    /// <summary>
    /// A command whose sole purpose is to relay its functionality to other objects by invoking delegates. The default return value for the CanExecute method is 'true'.
    /// </summary>
    public class RelayCommand : ICommand
    {
        /// <summary>
        /// Can execute function.
        /// </summary>
        private readonly Func<bool> canExecute;

        /// <summary>
        /// Execute command action.
        /// </summary>
        private readonly Action execute;

        /// <summary>
        /// Initializes a new instance of the <see cref="RelayCommand"/> class and the command can always be executed.
        /// </summary>
        /// <param name="execute">The execution logic.</param>
        public RelayCommand(Action execute)
            : this(execute, null)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RelayCommand"/> class.
        /// </summary>
        /// <param name="execute">The execution logic.</param>
        /// <param name="canExecute">The execution status logic.</param>
        public RelayCommand(Action execute, Func<bool> canExecute)
        {
            if (execute == null)
            {
                throw new ArgumentNullException("execute");
            }

            this.execute = execute;
            this.canExecute = canExecute;
        }

        public RelayCommand()
        {
            // TODO: Complete member initialization
        }

        /// <summary>
        /// Can execute event changed
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add
            {
                if (this.canExecute != null)
                {
                    CommandManager.RequerySuggested += value;
                }
            }

            remove
            {
                if (this.canExecute != null)
                {
                    CommandManager.RequerySuggested -= value;
                }
            }
        }

        /// <summary>
        /// Can execute command
        /// </summary>
        /// <param name="parameter">Attached parameter</param>
        /// <returns>Return true if can execute otherwise return false.</returns>
        public bool CanExecute(object parameter)
        {
            return this.canExecute == null ? true : this.canExecute();
        }

        /// <summary>
        /// Execute command.
        /// </summary>
        /// <param name="parameter">Attached parameter.</param>
        public void Execute(object parameter)
        {
            this.execute();
        }
    }
}
