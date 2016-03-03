using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPasswordRecovery.Core
{
    public class ExcelHashUtility
    {
        private static uint HashPassword(String password)
        {
            const uint magicConstant = 52811;

            List<uint> values = new List<uint>();

            for (int i = 0; i < password.Length; i++)
            {
                uint value = (uint)password[i] << i + 1;

                // works but not pretty
                while (value > 32767)
                {
                    value -= 32767;
                }

                values.Add(value);
            }

            var xorOfValues = values.Aggregate((x, y) => x ^ y);

            return xorOfValues ^ (uint)password.Length ^ magicConstant;
        }


        public static String FindPassword(string hash)
        {
            int passwordHashValue = int.Parse(hash, System.Globalization.NumberStyles.HexNumber);

            for (var i = 65; i <= 66; i++)
            {
                for (var j = 65; j <= 66; j++)
                {
                    for (var k = 65; k <= 66; k++)
                    {
                        for (var l = 65; l <= 66; l++)
                        {
                            for (var m = 65; m <= 66; m++)
                            {
                                for (var n = 65; n <= 66; n++)
                                {
                                    for (var i1 = 65; i1 <= 66; i1++)
                                    {
                                        for (var i2 = 65; i2 <= 66; i2++)
                                        {
                                            for (var i3 = 65; i3 <= 66; i3++)
                                            {
                                                for (var i4 = 65; i4 <= 66; i4++)
                                                {
                                                    for (var i5 = 65; i5 <= 66; i5++)
                                                    {
                                                        for (var i6 = 32; i6 <= 126; i6++)
                                                        {

                                                            string password = new string(new[]
                                                            {
                                                                (char)i, (char)j, (char)k, (char)l, (char)m, 
                                                                (char)i1,(char)i2,(char)i3,(char)i4,(char)i5,
                                                                (char)i6  
                                                            });

                                                            
                                                            if (passwordHashValue == HashPassword(password))
                                                            {
                                                                return password;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return null;
        }
    }
}
