using System;

namespace YanBo
{
    internal class Nf
    {
        private string zdnum;

        private string cbfmc;

        private string cbfdz;

        private string czs;

        private string zmj;

        private string[] hhMj;

        private string[] zmMj;

        private string[] qtMj;

        private string[] sub;

        private string year;

        public string getYear()
        {
            return this.year;
        }

        public void setYear(string year)
        {
            this.year = year;
        }

        public void setSubTotal(string[] sub)
        {
            this.sub = sub;
        }

        public string[] getSubTotal()
        {
            return this.sub;
        }

        public string getZdnum()
        {
            return this.zdnum;
        }

        public void setZdnum(string zdnum)
        {
            this.zdnum = zdnum;
        }

        public string getCbfmc()
        {
            return this.cbfmc;
        }

        public void setCbfmc(string cbfmc)
        {
            this.cbfmc = cbfmc;
        }

        public string getCbfdz()
        {
            return this.cbfdz;
        }

        public void setCbfdz(string cbfdz)
        {
            this.cbfdz = cbfdz;
        }

        public string getCzs()
        {
            return this.czs;
        }

        public void setCzs(string czs)
        {
            this.czs = czs;
        }

        public string getZmj()
        {
            return this.zmj;
        }

        public void setZmj(string zmj)
        {
            this.zmj = zmj;
        }

        public string[] getHhMj()
        {
            return this.hhMj;
        }

        public void setHhMj(string[] hhMj)
        {
            this.hhMj = hhMj;
        }

        public string[] getZmMj()
        {
            return this.zmMj;
        }

        public void setZmMj(string[] zmMj)
        {
            this.zmMj = zmMj;
        }

        public string[] getQtMj()
        {
            return this.qtMj;
        }

        public void setQtMj(string[] qtMj)
        {
            this.qtMj = qtMj;
        }

        public string toString()
        {
            return string.Concat(new string[]
			{
				"zdnum=",
				this.zdnum,
				", cbfmc=",
				this.cbfmc,
				", cbfdz=",
				this.cbfdz,
				", czs=",
				this.czs,
				", zmj=",
				this.zmj,
				", hhMj=",
				this.ArrayToString(this.hhMj),
				", zmMj=",
				this.ArrayToString(this.zmMj),
				", qtMj=",
				this.ArrayToString(this.qtMj),
				", subTotal=",
				this.SubToString(this.sub),
				", year=",
				this.year
			});
        }

        public string SubToString(string[] subTotal)
        {
            string result;
            if (subTotal == null)
            {
                result = "小计没有填写";
            }
            else
            {
                string text = "";
                for (int i = 0; i < subTotal.Length; i++)
                {
                    if (subTotal[i] != null && subTotal[i].Length >= 2)
                    {
                        if (i != 0 && text.Length > 2)
                        {
                            text += "_";
                        }
                        text += subTotal[i];
                    }
                }
                result = text;
            }
            return result;
        }

        public string ArrayToString(string[] array)
        {
            string text = "";
            string result;
            if (array == null)
            {
                result = text;
            }
            else
            {
                for (int i = 0; i < array.Length; i++)
                {
                    if (array[i] != null && !array[i].Trim().Equals(""))
                    {
                        if (i != 0 && text.Length > 0)
                        {
                            text += "_";
                        }
                        text += array[i];
                    }
                }
                result = text;
            }
            return result;
        }
    }
}
