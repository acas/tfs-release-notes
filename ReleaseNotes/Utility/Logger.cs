using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReleaseNotes
{
    class Logger
    {
        /// <summary>
        /// Describes the method for which a message is to be displayed
        /// </summary>
        public enum Method {
            Console,
            MessageBox
        }

        /// <summary>
        /// Describes message type
        /// </summary>
        public enum Type {
            Error,
            Warning,
            Information,
            Message,
            Success,
            General
        }

        private bool silence = false;
        private string logString = "";
        private Method method = Method.Console;
        private Type type = Type.Information;

        /// <summary>
        /// Empty logger constructor
        /// </summary>
        public Logger() {}

        /// <summary>
        /// Logger constructor with log string and ability to silence
        /// if necessary
        /// </summary>
        /// <param name="logString"></param>
        /// <param name="silence"></param>
        public Logger(string logString, bool silence)
        {
            this.logString = logString;
            this.silence = silence;
        }

        /// <summary>
        /// Logger constructor with added logging method selector
        /// </summary>
        /// <param name="logString"></param>
        /// <param name="silence"></param>
        /// <param name="method"></param>
        public Logger(string logString, bool silence, Method method)
        {
            this.logString = logString;
            this.silence = silence;
            this.method = method;
        }

        /// <summary>
        /// Logger constructor with added logging type selector
        /// </summary>
        /// <param name="logString"></param>
        /// <param name="silence"></param>
        /// <param name="type"></param>
        public Logger(string logString, bool silence, Type type)
        {
            this.logString = logString;
            this.silence = silence;
            this.type = type;
        }

        /// <summary>
        /// Constructor with both method and type selectors
        /// </summary>
        /// <param name="logString"></param>
        /// <param name="silence"></param>
        /// <param name="method"></param>
        /// <param name="type"></param>
        public Logger(string logString, bool silence, Method method, Type type)
        {
            this.logString = logString;
            this.silence = silence;
            this.method = method;
            this.type = type;
        }

        /// <summary>
        /// Displays the message currently in the logger object
        /// with the specified params
        /// </summary>
        /// <returns>A logger object</returns>
        public Logger display() 
        {
            MessageBoxIcon mbi = MessageBoxIcon.None;
            switch (this.type)
            {
                case Type.Error:
                    Console.ForegroundColor = ConsoleColor.Red;
                    mbi = MessageBoxIcon.Error;
                    break;
                case Type.Information:
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    mbi = MessageBoxIcon.Asterisk;
                    break;
                case Type.Warning:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    mbi = MessageBoxIcon.Warning;
                    break;
                case Type.Message:
                    Console.ForegroundColor = ConsoleColor.White;
                    mbi = MessageBoxIcon.None;
                    break;
                case Type.Success:
                    Console.ForegroundColor = ConsoleColor.Green;
                    mbi = MessageBoxIcon.Exclamation;
                    break;
                case Type.General:
                    Console.ForegroundColor = ConsoleColor.White;
                    mbi = MessageBoxIcon.None;
                    break;
            }

            switch (this.method) 
            {
                
                case Method.Console:
                    if (!this.silence)
                        if (this.type != Type.General)
                            Console.WriteLine(" [" + this.type.ToString() + ": " + this.logString + " @ " + DateTime.Now.ToShortTimeString() + " ]");
                        else
                            Console.WriteLine(" [ " + this.logString + " ]");
                    break;
                
                case Method.MessageBox:
                    if (!this.silence)
                        MessageBox.Show(null, logString, this.type.ToString(), MessageBoxButtons.OK, mbi);
                    break;
            }

            return this;
        }

        /// <summary>
        /// Sets the logging message
        /// </summary>
        /// <param name="message"></param>
        /// <returns>A logger object</returns>
        public Logger setMessage(string message)
        {
            this.logString = message;
            return this;
        }

        /// <summary>
        /// Sets the logging type
        /// </summary>
        /// <param name="type"></param>
        /// <returns>A logger object</returns>
        public Logger setType(Type type)
        {
            this.type = type;
            return this;
        }

        /// <summary>
        /// Sets the logging method
        /// </summary>
        /// <param name="method"></param>
        /// <returns>A logger object</returns>
        public Logger setMethod(Method method)
        {
            this.method = method;
            return this;
        }

        /// <summary>
        /// Sets whether or not the logged message will appear on the screen
        /// </summary>
        /// <param name="silent"></param>
        /// <returns>A logger object</returns>
        public Logger setSilence(bool silent)
        {
            this.silence = silent;
            return this;
        }
    }
}
