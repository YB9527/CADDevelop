using System;
using System.Collections.Generic;

namespace YanBo
{
    internal class Jzx
    {
        private string cbfmc;

        private string cbfbm;

        private string cbfdz;

        private List<string> jzxSm;

        private List<string> jzxBs;

        public Jzx(string cbfdz)
        {
            this.cbfdz = cbfdz;
        }

        public string getCbfmc()
        {
            return this.cbfmc;
        }

        public void setCbfmc(string cbfmc)
        {
            this.cbfmc = cbfmc;
        }

        public string getCbfbm()
        {
            return this.cbfbm;
        }

        public void setCbfbm(string cbfbm)
        {
            this.cbfbm = cbfbm;
        }

        public string getCbfdz()
        {
            return this.cbfdz;
        }

        public void setCbfdz(string cbfdz)
        {
            this.cbfdz = cbfdz;
        }

        public List<string> getJzxSm()
        {
            return this.jzxSm;
        }

        public void setJzxSm(List<string> jzxSm)
        {
            this.jzxSm = jzxSm;
        }

        public List<string> getJzxBs()
        {
            return this.jzxBs;
        }

        public void setJzxBs(List<string> jzxBs)
        {
            this.jzxBs = jzxBs;
        }
    }
}
