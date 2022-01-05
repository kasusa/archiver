using System;
using System.Text;

namespace archiver.ConsoleColorWriter
{
    /// <summary>
    /// Very easy to use console colors.
    /// Works out of the box, no initializing necessary. Just call the static methods everything is setup already.
    /// You can of course change all the colors being used by the methods by changing the properties ( ex: StandardBackColor ).
    /// Made by SHEePYTaGGeRNeP
    /// </summary>
    public static class ConsoleWriter
    {
        #region - - - - - - - - - - - - - - - - - - Color Properties - - - - - - - - - - - - - - - - - - 

        /// <summary>
        /// Standard Front color. Used in WriteSeperator and WriteText
        /// </summary>
        public static ConsoleColor StandardFrontColor { get; set; } = ConsoleColor.White;

        /// <summary>
        /// Standard Background color. Used in WriteSeperator and WriteText
        /// </summary>
        public static ConsoleColor StandardBackColor { get; set; } = ConsoleColor.Black;

        /// <summary>
        /// Used in WriteQuestionMessage
        /// </summary>
        public static ConsoleColor QuestionFrontColor { get; set; } = ConsoleColor.Black;

        /// <summary>
        /// Used in WriteQuestionMessage
        /// </summary>
        public static ConsoleColor QuestionBackColor { get; set; } = ConsoleColor.White;

        /// <summary>
        /// Used in WriteErrorMessage. This is changes the "ERROR: " color. 
        /// </summary>
        public static ConsoleColor ErrorFrontColor { get; set; } = ConsoleColor.Red;

        /// <summary>
        /// Used in WriteErrorMessage. This is changes the "ERROR: " color. 
        /// </summary>
        public static ConsoleColor ErrorBackColor { get; set; } = ConsoleColor.Black;

        /// <summary>
        /// Used in WriteErrorMessage. This changes the color of the actual message.
        /// </summary>
        public static ConsoleColor ErrorMessageFrontColor { get; set; } = ConsoleColor.White;

        /// <summary>
        /// Used in WriteErrorMessage. This changes the color of the actual message.
        /// </summary>
        public static ConsoleColor ErrorMessageBackColor { get; set; } = ConsoleColor.Red;

        /// <summary>
        /// Used in WriteWarningMessage. This is changes the "WARNING: " color.
        /// </summary>
        public static ConsoleColor WarningFrontColor { get; set; } = ConsoleColor.DarkYellow;

        /// <summary>
        /// Used in WriteWarningMessage. This is changes the "WARNING: " color. 
        /// </summary>
        public static ConsoleColor WarningBackColor { get; set; } = ConsoleColor.Black;

        /// <summary>
        /// Used in WriteWarningMessage. This changes the color of the actual message.
        /// </summary>
        public static ConsoleColor WarningMessageFrontColor { get; set; } = ConsoleColor.White;

        /// <summary>
        /// Used in WriteWarningMessage. This changes the color of the actual message.
        /// </summary>
        public static ConsoleColor WarningMessageBackColor { get; set; } = ConsoleColor.DarkYellow;

        #endregion

        /// <summary>
        /// Sets the current console colors.
        /// </summary>
        public static void SetColors(ConsoleColor pFrontColor, ConsoleColor pBackgroundColor)
        {
            Console.ForegroundColor = pFrontColor;
            Console.BackgroundColor = pBackgroundColor;
        }

        /// <summary>
        /// Writes a new line full of [pCharacter] in the CURRENT CONSOLE COLORS.
        /// </summary>
        public static void WriteSeperator(char pCharacter)
        {
            StringBuilder lvMBuilder = new StringBuilder();
            for (int lvIndex = 0; lvIndex < Console.BufferWidth; lvIndex++)
                lvMBuilder.Append(pCharacter);
            Console.WriteLine(lvMBuilder.ToString());
        }

        /// <summary>
        /// Writes with the StandardFrontColor and StandardBackColor, creates a new line and then resets the console colors to StandardFrontColor and StandardBackColor.
        /// </summary>
        /// <param name="pText">Text to be written to console</param>
        /// <param name="pWriteLine">TRUE = creates a new line after text has been placed.</param>
        /// <param name="pWriteTime">TRUE = writes the current time in StandardFrontColor and StandardBackColor before writing the text.</param>
        public static void WriteText(string pText, bool pWriteLine = true, bool pWriteTime = true)
        {
            WriteColoredText(pText, StandardFrontColor, StandardBackColor, pWriteLine, pWriteTime);
        }

        /// <summary>
        /// Writes with the QuestionFrontColor and QuestionBackColor and then resets the console colors to StandardFrontColor and StandardBackColor.
        /// </summary>
        /// <param name="pText">Text to be written to console</param>
        /// <param name="pWriteLine">TRUE = creates a new line after text has been placed.</param>
        /// <param name="pWriteTime">TRUE = writes the current time in StandardFrontColor and StandardBackColor before writing the text.</param>
        public static void WriteQuestionMessage(string pText, bool pWriteLine = true, bool pWriteTime = true)
        {
            WriteColoredText(pText, QuestionFrontColor, QuestionBackColor, pWriteLine, pWriteTime);
        }

        /// <summary>
        /// Writes with the ErrorFrontColor and ErrorBackColor and then resets the console colors to StandardFrontColor and StandardBackColor.
        /// </summary>
        /// <param name="pText">Text to be written to console.</param>
        /// <param name="pWriteLine">TRUE = creates a new line after text has been placed.</param>
        /// <param name="pWriteTime">TRUE = writes the current time in StandardFrontColor and StandardBackColor before writing the text.</param>
        public static void WriteErrorMessage(string pText, bool pWriteLine = true, bool pWriteTime = true)
        {
            WriteColoredText("ERROR: ", ErrorFrontColor, ErrorBackColor, false, pWriteTime);
            WriteColoredText(pText, ErrorMessageFrontColor, ErrorMessageBackColor, pWriteLine, false);
        }

        /// <summary>
        /// Writes with the WarningFrontColor and WarningBackColor and then resets the console colors to StandardFrontColor and StandardBackColor.
        /// </summary>
        /// <param name="pText">Text to be written to console.</param>
        /// <param name="pWriteLine">TRUE = creates a new line after text has been placed.</param>
        /// <param name="pWriteTime">TRUE = writes the current time in StandardFrontColor and StandardBackColor before writing the text.</param>
        public static void WriteWarningMessage(string pText, bool pWriteLine = true, bool pWriteTime = true)
        {
            WriteColoredText("WARNING: ", WarningFrontColor, WarningBackColor, false, pWriteTime);
            WriteColoredText(pText, WarningMessageFrontColor, WarningMessageBackColor, true, false);
        }


        /// <summary>
        /// Writes Colored text to the console, uses StandardBackColor and then resets the console colors to StandardFrontColor and StandardBackColor.
        /// </summary>
        /// <param name="pText">Text to be written to console.</param>
        /// <param name="pFrontColor">The front color</param>
        /// <param name="pWriteLine">TRUE = creates a new line after text has been placed.</param>
        /// <param name="pWriteCurrentTime">TRUE = writes the current time in StandardFrontColor and StandardBackColor before writing the text.</param>
        public static void WriteColoredText(string pText, ConsoleColor pFrontColor, bool pWriteLine = true, bool pWriteCurrentTime = false)
        {
            WriteColoredText(pText, pFrontColor, StandardBackColor, pWriteLine, pWriteCurrentTime);
        }
        /// <summary>
        /// Writes Colored text to the console and then resets the console colors to StandardFrontColor and StandardBackColor.
        /// </summary>
        /// <param name="pText">Text to be written to console.</param>
        /// <param name="pFrontColor">The front color</param>
        /// <param name="pBackgroundColor">The back color</param>
        /// <param name="pWriteLine">TRUE = creates a new line after text has been placed.</param>
        /// <param name="pWriteCurrentTime">TRUE = writes the current time in StandardFrontColor and StandardBackColor before writing the text.</param>
        public static void WriteColoredText(string pText, ConsoleColor pFrontColor, ConsoleColor pBackgroundColor, bool pWriteLine = true, bool pWriteCurrentTime = true)
        {
            if (pWriteCurrentTime)
            {
                Console.ForegroundColor = StandardFrontColor;
                Console.BackgroundColor = StandardBackColor;
                Console.Write(DateTime.Now + " : ");
            }
            Console.ForegroundColor = pFrontColor;
            Console.BackgroundColor = pBackgroundColor;
            if (pWriteLine)
                Console.WriteLine(pText);
            else
                Console.Write(pText);
            Console.ForegroundColor = StandardFrontColor;
            Console.BackgroundColor = StandardBackColor;
        }
    }
}