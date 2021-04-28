using System;
using System.Windows;
using System.Windows.Input;

namespace TinhThuong
{
    /// <summary>
    /// Defines the attached properties to create a CommandBehaviorBinding
    /// </summary>
    public class CommandBehavior
    {
        #region Behavior

        /// <summary>
        /// Behavior Attached Dependency Property
        /// </summary>
        private static readonly DependencyProperty BehaviorProperty =
            DependencyProperty.RegisterAttached("Behavior", typeof(CommandBehaviorBinding), typeof(CommandBehavior),
                new FrameworkPropertyMetadata((CommandBehaviorBinding)null));

        /// <summary>
        /// Gets the Behavior property. 
        /// </summary>
        private static CommandBehaviorBinding GetBehavior(DependencyObject d)
        {
            return (CommandBehaviorBinding)d.GetValue(BehaviorProperty);
        }

        /// <summary>
        /// Sets the Behavior property.  
        /// </summary>
        private static void SetBehavior(DependencyObject d, CommandBehaviorBinding value)
        {
            d.SetValue(BehaviorProperty, value);
        }

        #endregion

        #region Command

        /// <summary>
        /// Command Attached Dependency Property
        /// </summary>
        public static readonly DependencyProperty CommandProperty =
            DependencyProperty.RegisterAttached("Command", typeof(ICommand), typeof(CommandBehavior),
                new FrameworkPropertyMetadata((ICommand)null,
                    new PropertyChangedCallback(OnCommandChanged)));

        /// <summary>
        /// Gets the Command property.  
        /// </summary>
        public static ICommand GetCommand(DependencyObject d)
        {
            return (ICommand)d.GetValue(CommandProperty);
        }

        /// <summary>
        /// Sets the Command property. 
        /// </summary>
        public static void SetCommand(DependencyObject d, ICommand value)
        {
            d.SetValue(CommandProperty, value);
        }

        /// <summary>
        /// Handles changes to the Command property.
        /// </summary>
        private static void OnCommandChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CommandBehaviorBinding binding = FetchOrCreateBinding(d);
            binding.Command = (ICommand)e.NewValue;
        }

        #endregion

        #region Action

        /// <summary>
        /// Action Attached Dependency Property
        /// </summary>
        public static readonly DependencyProperty ActionProperty =
            DependencyProperty.RegisterAttached("Action", typeof(Action<object>), typeof(CommandBehavior),
                new FrameworkPropertyMetadata((Action<object>)null,
                    new PropertyChangedCallback(OnActionChanged)));

        /// <summary>
        /// Gets the Action property.  
        /// </summary>
        public static Action<object> GetAction(DependencyObject d)
        {
            return (Action<object>)d.GetValue(ActionProperty);
        }

        /// <summary>
        /// Sets the Action property. 
        /// </summary>
        public static void SetAction(DependencyObject d, Action<object> value)
        {
            d.SetValue(ActionProperty, value);
        }

        /// <summary>
        /// Handles changes to the Action property.
        /// </summary>
        private static void OnActionChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CommandBehaviorBinding binding = FetchOrCreateBinding(d);
            binding.Action = (Action<object>)e.NewValue;
        }

        #endregion
        #region Event Mouse

       /* public ICommand SelectedItemDoubleClickedCommand
        {
            get { return (ICommand)GetValue(SelectedItemDoubleClickedCommandProperty); }
            set { SetValue(SelectedItemDoubleClickedCommandProperty, value); }
        }
        public static readonly DependencyProperty SelectedItemDoubleClickedCommandProperty = DependencyProperty.Register(
            "SelectedItemDoubleClickedCommand", typeof(ICommand),
            typeof(CommandBehavior),
            new UIPropertyMetadata(null));
        public static ICommand SelectedItemDoubleClickedCommandAttached(DependencyObject obj)
        {
            return (ICommand)obj.GetValue(SelectedItemDoubleClickedCommandProperty);
        }
        public static void SetSelectedItemDoubleClickedCommandAttached(DependencyObject obj, ICommand value)
        {
            obj.SetValue(SelectedItemDoubleClickedCommandAttachedProperty, value);
        }
        public static readonly DependencyProperty SelectedItemDoubleClickedCommandAttachedProperty = DependencyProperty.RegisterAttached(
            "SelectedItemDoubleClickedCommandAttached",
            typeof(ICommand), typeof(CommandBehavior),
            new UIPropertyMetadata(null, SelectedItemDoubleClickedCommandAttachedChanged));

        private static void SelectedItemDoubleClickedCommandAttachedChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var item = d as TreeViewItem;
            if (item != null)
            {
                if (e.NewValue != null)
                {
                    var binding = new MouseBinding((ICommand)e.NewValue, new MouseGesture(MouseAction.LeftDoubleClick));

                    BindingOperations.SetBinding(binding, InputBinding.CommandParameterProperty, new Binding("SelectedItem")
                    {
                        RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(TreeView), 1)
                    });

                    item.InputBindings.Add(binding);
                }
            }
        }*/
        #endregion

        #region CommandParameter

        /// <summary>
        /// CommandParameter Attached Dependency Property
        /// </summary>
        public static readonly DependencyProperty CommandParameterProperty =
            DependencyProperty.RegisterAttached("CommandParameter", typeof(object), typeof(CommandBehavior),
                new FrameworkPropertyMetadata((object)null,
                    new PropertyChangedCallback(OnCommandParameterChanged)));

        /// <summary>
        /// Gets the CommandParameter property.  
        /// </summary>
        public static object GetCommandParameter(DependencyObject d)
        {
            return (object)d.GetValue(CommandParameterProperty);
        }

        /// <summary>
        /// Sets the CommandParameter property. 
        /// </summary>
        public static void SetCommandParameter(DependencyObject d, object value)
        {
            d.SetValue(CommandParameterProperty, value);
        }

        /// <summary>
        /// Handles changes to the CommandParameter property.
        /// </summary>
        private static void OnCommandParameterChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CommandBehaviorBinding binding = FetchOrCreateBinding(d);
            binding.CommandParameter = e.NewValue;
        }

        #endregion

        #region Event

        /// <summary>
        /// Event Attached Dependency Property
        /// </summary>
        public static readonly DependencyProperty EventProperty =
            DependencyProperty.RegisterAttached("Event", typeof(string), typeof(CommandBehavior),
                new FrameworkPropertyMetadata((string)String.Empty,
                    new PropertyChangedCallback(OnEventChanged)));

        /// <summary>
        /// Gets the Event property.  This dependency property 
        /// indicates ....
        /// </summary>
        public static string GetEvent(DependencyObject d)
        {
            return (string)d.GetValue(EventProperty);
        }

        /// <summary>
        /// Sets the Event property.  This dependency property 
        /// indicates ....
        /// </summary>
        public static void SetEvent(DependencyObject d, string value)
        {
            d.SetValue(EventProperty, value);
        }

        /// <summary>
        /// Handles changes to the Event property.
        /// </summary>
        private static void OnEventChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CommandBehaviorBinding binding = FetchOrCreateBinding(d);
            //check if the Event is set. If yes we need to rebind the Command to the new event and unregister the old one
            if (binding.Event != null && binding.Owner != null)
                binding.Dispose();
            //bind the new event to the command
            binding.BindEvent(d, e.NewValue.ToString());
        }

        #endregion

        #region Helpers
        //tries to get a CommandBehaviorBinding from the element. Creates a new instance if there is not one attached
        private static CommandBehaviorBinding FetchOrCreateBinding(DependencyObject d)
        {
            CommandBehaviorBinding binding = CommandBehavior.GetBehavior(d);
            if (binding == null)
            {
                binding = new CommandBehaviorBinding();
                CommandBehavior.SetBehavior(d, binding);
            }
            return binding;
        }
        #endregion

    }

}
