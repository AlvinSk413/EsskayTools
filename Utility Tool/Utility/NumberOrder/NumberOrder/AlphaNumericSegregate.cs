using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NumberOrder
{
    class AlphaNumericSegregate
    {
        public string alpha { get; set; }
        public string number { get; set; }
        public string value { get; set; }
        public List<AlphaNumericSegregate> alphaNumericSegregates { get; set; }
        public AlphaNumericSegregate (string[] alpha)
        {
            List<string> alphaNumeric = new List<string>();
            foreach(string value in alpha)
            {
                bool flag = alphaNumeric.Any(x => x.Contains(value));
                if(flag==false)
                {
                    alphaNumeric.Add(value);
                }
            }
            List<AlphaNumericSegregate> alphaNumericSegregates = new List<AlphaNumericSegregate>();
            foreach(string alp in alphaNumeric)
            {
                AlphaNumericSegregate alphaNumericSegregate = new AlphaNumericSegregate();
                string numberConcate = "";
                string alphaConcate ="";
                foreach (char charr in alp)
                {
                    
                    bool flag = char.IsLetterOrDigit(charr);
                    if(flag)
                    {
                        bool letter = char.IsLetter(charr);
                        if (letter)
                        {
                            alphaConcate = alphaConcate + charr;
                            
                        }
                        else
                        {
                            numberConcate = numberConcate + charr;
                            
                        }

                    }
                   
                    
                }
                alphaNumericSegregate.alpha = alphaConcate;
                alphaNumericSegregate.number = numberConcate;
                alphaNumericSegregate.value = alp;
                alphaNumericSegregates.Add(alphaNumericSegregate);

            }
            this.alphaNumericSegregates = alphaNumericSegregates;
            
        }
        public AlphaNumericSegregate  ()
        {

        }
    }
}
