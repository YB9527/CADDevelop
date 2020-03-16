using System;
using System.Collections.Generic;
using System.Text;

namespace YanBo
{
    class TreeSet<T> : List<T>
    {
        public new  void Add(T t)
        {
            if(!this.Contains(t))
            {
                this.Add(t);
            }
        }
    }
}
