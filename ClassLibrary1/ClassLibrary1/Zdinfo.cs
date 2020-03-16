using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using YanBo.ZJDPo;

namespace YanBo
{
    internal class Zdinfo
    {
        private string zdnum;

        private string tdlb;

        private string quanli;

        private string qh;

        private string tufu;

        private string tdsyz;

        private string txdz;

        private string tdzl;

        private string dz;

        private string nz;

        private string xz;

        private string bz;

        private string sbjzwqs;

        private string shrq;

        private string djrq;

        public Polyline Polyline;

        public Point3dCollection PC;

        public string jzmj;

        public string area;

        public ZDX zdx;

        public List<Jzxinfo> JZXS;

        public string getZdnum()
        {
            return this.zdnum;
        }

        public void setZdnum(string zdnum)
        {
            this.zdnum = zdnum;
        }

        public string getTdlb()
        {
            return this.tdlb;
        }

        public void setTdlb(string tdlb)
        {
            this.tdlb = tdlb;
        }

        public string getQuanli()
        {
            return this.quanli;
        }

        public void setQuanli(string quanli)
        {
            this.quanli = quanli;
        }

        public string getQh()
        {
            return this.qh;
        }

        public void setQh(string qh)
        {
            this.qh = qh;
        }

        public string getTufu()
        {
            return this.tufu;
        }

        public void setTufu(string tufu)
        {
            this.tufu = tufu;
        }

        public string getTdsyz()
        {
            return this.tdsyz;
        }

        public void setTdsyz(string tdsyz)
        {
            this.tdsyz = tdsyz;
        }

        public string getTxdz()
        {
            return this.txdz;
        }

        public void setTxdz(string txdz)
        {
            this.txdz = txdz;
        }

        public string getTdzl()
        {
            return this.tdzl;
        }

        public void setTdzl(string tdzl)
        {
            this.tdzl = tdzl;
        }

        public string getDz()
        {
            return this.dz;
        }

        public void setDz(string dz)
        {
            this.dz = dz;
        }

        public string getNz()
        {
            return this.nz;
        }

        public void setNz(string nz)
        {
            this.nz = nz;
        }

        public string getXz()
        {
            return this.xz;
        }

        public void setXz(string xz)
        {
            this.xz = xz;
        }

        public string getBz()
        {
            return this.bz;
        }

        public void setBz(string bz)
        {
            this.bz = bz;
        }

        public string getSbjzwqs()
        {
            return this.sbjzwqs;
        }

        public void setSbjzwqs(string sbjzwqs)
        {
            this.sbjzwqs = sbjzwqs;
        }

        public string getShrq()
        {
            return this.shrq;
        }

        public void setShrq(string shrq)
        {
            this.shrq = shrq;
        }

        public string getDjrq()
        {
            return this.djrq;
        }

        public void setDjrq(string djrq)
        {
            this.djrq = djrq;
        }
    }
}
