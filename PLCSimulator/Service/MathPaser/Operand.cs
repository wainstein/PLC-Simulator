using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PLCTools.Service.MathPaser
{
    /// Numerical Operation type	
    public enum OperandType
    {
        FUNC = 1,
        DATE = 2,
        NUMBER = 3,
        BOOLEAN = 4,
        STRING = 5

    }

	/// Operand types(Weight going down from up to bottom)，the greater of the number，the lower it has the priority.		
	public enum OperatorType
	{
		/// (,left bracket	
		LB = 10,
		/// ),right bracket	
		RB = 11,
		/// !,NOT	
		NOT = 20,
		/// +,positive sign	
		PS = 21,
		/// -,negative sign	
		NS = 22,
		/// tan	
		TAN = 23,
		/// atan
		ATAN = 24,
        /// ^,power calculation
        PWR = 25,
		/// *,multiplication
		MUL = 30,
		/// /,division	
        DIV = 31,
		/// %,modulus	
        MOD = 32,
		/// +,Addition
		ADD = 40,
		/// -,subtraction
		SUB = 41,
		/// less than
		LT = 50,
		/// less than or equal to	
		LE = 51,
		/// >,greater than	
		GT = 52,	
        /// >=,greater than or equal to	
        GE = 53,
        /// =,equal to	
		ET = 60,
		/// unequal to
		UT = 61,
		/// &,AND
		AND = 70,
		/// |,OR
		OR = 71,
		///,comma
		CA = 80,
		/// ending @	
		END = 255,
        /// Error
        ERR = 256
	}


	public class Operand
    {
        #region Constructed Function			
        public Operand(OperandType type, object value)
        {
            this.Type = type;
            this.Value = value;
        }

        public Operand(string opd, object value)
        {
            this.Type = ConvertOperand(opd);
            this.Value = value;
        }
        #endregion

        #region Variable &　Property			
        /// Operation type			
        public OperandType Type { get; set; }
        /// Keyword	
        public string Key { get; set; }
        /// Operate value			
        public object Value { get; set; }

        #endregion

        #region Public Method			
        /// Convert the Operand to specified type
        /// <param name="opd">operand</param>			
        /// <returns>Return related operand type</returns>			
        public static OperandType ConvertOperand(string opd)
        {
            if (opd.IndexOf("(") > -1)
            {
                return OperandType.FUNC;
            }
            else if (IsNumber(opd))
            {
                return OperandType.NUMBER;
            }
            else if (IsDate(opd))
            {
                return OperandType.DATE;
            }
            else
            {
                return OperandType.STRING;
            }
        }

        /// <summary>			
        /// check if the object number			
        /// </summary>			
        /// <param name="value">Object value</param>			
        /// <returns>if yes return true, otherwise return false</returns>			
        public static bool IsNumber(object value)
        {
            double val;
            return double.TryParse(value.ToString(), out val);
        }

        /// <summary>			
        /// check if the object date			
        /// </summary>			
        /// <param name="value">Object value</param>			
        /// <returns>if yes return true, otherwise return false</returns>			
        public static bool IsDate(object value)
        {
            DateTime dt;
            return DateTime.TryParse(value.ToString(), out dt);
        }
        #endregion
    }

}
