using System;

namespace YanBo
{
    internal class Qzb
    {
        private string bzdh;

        private string lzdh;

        private string qdh;

        private string zdh;

        private string lzdzjr;

        public string toString()
        {
            return string.Concat(new string[]
			{
				"A [bzdh=",
				this.bzdh,
				", lzdh=",
				this.lzdh,
				", qdh=",
				this.qdh,
				", zdh=",
				this.zdh,
				", lzdzjr=",
				this.lzdzjr,
				"]"
			});
        }

        public string getBzdh()
        {
            return this.bzdh;
        }

        public void setBzdh(string bzdh)
        {
            this.bzdh = bzdh;
        }

        public string getLzdh()
        {
            return this.lzdh;
        }

        public void setLzdh(string lzdh)
        {
            this.lzdh = lzdh;
        }

        public string getQdh()
        {
            return this.qdh;
        }

        public void setQdh(string qdh)
        {
            this.qdh = qdh;
        }

        public string getZdh()
        {
            return this.zdh;
        }

        public void setZdh(string zdh)
        {
            this.zdh = zdh;
        }

        public string getLzdzjr()
        {
            return this.lzdzjr;
        }

        public void setLzdzjr(string lzdzjr)
        {
            this.lzdzjr = lzdzjr;
        }
    }
}
