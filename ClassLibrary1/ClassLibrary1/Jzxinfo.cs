using System;

namespace YanBo
{
    internal class Jzxinfo
    {
        private string bzdh;

        private string lzdh;

        private string qdh;

        private string zdh;

        private double tsbc;

        private double kzbc;

        private string jxxz;

        private string jzxlb;

        private string jzxwz;

        private string bzdzjr;

        private string bzdzjrq;

        private string lzdzjr;

        private string lzdzjrq;

        private string qdzb;

        private string zdzb;

        public string toString()
        {
            return string.Concat(new object[]
			{
				"A [bzdh=",
				this.bzdh,
				", lzdh=",
				this.lzdh,
				", qdh=",
				this.qdh,
				", zdh=",
				this.zdh,
				", tsbc=",
				this.tsbc,
				", kzbc=",
				this.kzbc,
				", jxxz=",
				this.jxxz,
				", jzxwz=",
				this.jzxwz,
				", bzdzjr=",
				this.bzdzjr,
				", bzdzjrq=",
				this.bzdzjrq,
				", lzdzjr=",
				this.lzdzjr,
				", lzdzjrq=",
				this.lzdzjrq,
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

        public double getTsbc()
        {
            return this.tsbc;
        }

        public void setTsbc(double tsbc)
        {
            this.tsbc = tsbc;
        }

        public double getKzbc()
        {
            return this.kzbc;
        }

        public void setKzbc(double kzbc)
        {
            this.kzbc = kzbc;
        }

        public string getJxxz()
        {
            return this.jxxz;
        }

        public void setJxxz(string jxxz)
        {
            this.jxxz = jxxz;
        }

        public string getJzxwz()
        {
            return this.jzxwz;
        }

        public void setJzxwz(string jzxwz)
        {
            this.jzxwz = jzxwz;
        }

        public string getBzdzjr()
        {
            return this.bzdzjr;
        }

        public void setBzdzjr(string bzdzjr)
        {
            this.bzdzjr = bzdzjr;
        }

        public string getBzdzjrq()
        {
            return this.bzdzjrq;
        }

        public void setBzdzjrq(string bzdzjrq)
        {
            this.bzdzjrq = bzdzjrq;
        }

        public string getLzdzjr()
        {
            return this.lzdzjr;
        }

        public void setLzdzjr(string lzdzjr)
        {
            this.lzdzjr = lzdzjr;
        }

        public string getLzdzjrq()
        {
            return this.lzdzjrq;
        }

        public void setLzdzjrq(string lzdzjrq)
        {
            this.lzdzjrq = lzdzjrq;
        }

        public string getJzxlb()
        {
            return this.jzxlb;
        }

        public void setJzxlb(string jzxlb)
        {
            this.jzxlb = jzxlb;
        }

        public string getQdzb()
        {
            return this.qdzb;
        }

        public void setQdzb(string qdzb)
        {
            this.qdzb = qdzb;
        }

        public string getZdzb()
        {
            return this.zdzb;
        }

        public void setZdzb(string zdzb)
        {
            this.zdzb = zdzb;
        }
    }
}
