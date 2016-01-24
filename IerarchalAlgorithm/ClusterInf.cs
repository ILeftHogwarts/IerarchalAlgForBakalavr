using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IerarchalAlgorithm
{
    class ClusterInf : ICloneable
    {
        public List<double[]> Cluster { set; get; }

        public ClusterInf()
        {
            Cluster = new List<double[]>();
        }

        public object Clone()
        {
            ClusterInf newCluster = new ClusterInf {Cluster = new List<double[]>(Cluster)};
         
            return newCluster;
        }

        public bool HasElem(double[] point)
        {
            bool flag = true;
            foreach (double[] chekblPoint in Cluster)
            {
                flag = true;
                for (int i = 0; i < point.Length; i++)
                {
                    
                    if (chekblPoint[i]!=point[i])
                    flag = false;
                }
                if (flag)
                {
                    return flag;
                }
                
            }
            return flag;
        }
         
    }
}
