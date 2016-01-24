using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;


namespace IerarchalAlgorithm
{
    public partial class Form1 : Form
    {
        private delegate double[] FindOpDeleg(double[] x, double[] y);

        private static List<ClusterInf> upDataList;
        private static List<ClusterInf> downDataList;
        public Form1()
        {
            InitializeComponent();
            upDataList = GetDataForUpIerAlg();
            downDataList = GetDataToDownIerarAlg();
        }

        private void AlgorithmCalcButton_Click(object sender, EventArgs e)
        {
            UpIerarch.Series.Clear();
            upDataList = UpIerarchAlg(upDataList);
            int i = 0;
            foreach (ClusterInf cluster in upDataList)
            {
                UpIerarch.Series.Add(i.ToString());
                UpIerarch.Series[i].ChartType = SeriesChartType.Point;
                UpIerarch.Series[i].MarkerSize = 10;
                foreach (double[] point in cluster.Cluster)
                {
                    UpIerarch.Series[i].Points.AddXY(point[0],point[1]);
                }
                i++;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UpIerarch.Series.Clear();
            downDataList = DownIerarchAlg(downDataList);
            int i = 0;
            foreach (ClusterInf cluster in downDataList)
            {
                UpIerarch.Series.Add(i.ToString());
                UpIerarch.Series[i].ChartType = SeriesChartType.Point;
                UpIerarch.Series[i].MarkerSize = 10;
                foreach (double[] point in cluster.Cluster)
                {
                    UpIerarch.Series[i].Points.AddXY(point[0], point[1]);
                }
                i++;
            }
        }

        private List<ClusterInf> GetDataForUpIerAlg()
        {
            List<ClusterInf> dataList = new List<ClusterInf>();
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"C:\Универ\Диплом\IerarchalAlgorithm\data.xlsx",
                Type.Missing,true);
            Excel.Worksheet objWorksheet = (Excel.Worksheet) ObjWorkBook.Sheets[1];
            Excel.Range excelRange = objWorksheet.UsedRange;
            
            for (int i = 0; i < excelRange.Rows.Count; i++)
            {
                double[] point = new double[excelRange.Columns.Count];
                for (int j = 0; j < excelRange.Columns.Count; j++)
                {
                        point[j] = Convert.ToDouble(objWorksheet.Cells[i + 1, j + 1].Text);
                }
                ClusterInf clust = new ClusterInf();
                clust.Cluster = new List<double[]> {point};
                dataList.Add(clust);
            }
            ObjWorkBook.Close(false, null, null);
            ObjWorkExcel.Quit();
            return dataList;
        }

        private List<ClusterInf> GetDataToDownIerarAlg()
        {
            List<ClusterInf> dataList = new List<ClusterInf>();
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"C:\Универ\Диплом\IerarchalAlgorithm\data.xlsx",
                Type.Missing, true);
            Excel.Worksheet objWorksheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            Excel.Range excelRange = objWorksheet.UsedRange;
            ClusterInf clust = new ClusterInf();
            clust.Cluster = new List<double[]>();
            for (int i = 0; i < excelRange.Rows.Count; i++)
            {
                double[] point = new double[excelRange.Columns.Count];
                for (int j = 0; j < excelRange.Columns.Count; j++)
                {
                    point[j] = Convert.ToDouble(objWorksheet.Cells[i + 1, j + 1].Text);
                }
                clust.Cluster.Add(point);
            }
            dataList.Add(clust);
            ObjWorkBook.Close(false, null, null);
            ObjWorkExcel.Quit();
            return dataList;
        } 

        private List<ClusterInf> UpIerarchAlg(List<ClusterInf> dataList)
        {
            
            double dist = 0;
            double minDist = Double.MaxValue;
            //ClusterInf minNextClus = null;
            int indPrevClus = -1, indNextClus = -1;
            foreach (ClusterInf prevCluster in dataList)
            {
                foreach (ClusterInf nextCluster in dataList)
                {
                    if (prevCluster==nextCluster)
                        continue;
                    foreach (double[] elInPrevClus in prevCluster.Cluster)
                    {
                       foreach (double[] elInNextClus in nextCluster.Cluster)
                       {
                            dist = 0;
                            for (int i = 0; i < elInPrevClus.Length; i++) // алгоритм поиска расстояния ближнего соседа 
                            {
                                dist += Math.Pow((elInPrevClus[i] - elInNextClus[i]), 2); 
                            }
                            if (dist < minDist)
                            {
                                minDist = dist;
                                indNextClus = dataList.IndexOf(prevCluster);
                                indPrevClus = dataList.IndexOf(nextCluster);
                                
                            }
                            
                        }
                        //double[] fPoint = FindInDeep(elInPrevClus, nextCluster.Cluster, minDist);
                    }
                }
            }
            foreach (double[] point in dataList[indNextClus].Cluster)
            {
                dataList[indPrevClus].Cluster.Add(point);
            }
            dataList.RemoveAt(indNextClus);
            return dataList;
        }

        private List<ClusterInf> DownIerarchAlg(List<ClusterInf> dataList)
        {
            double dist = 0, maxDist = 0, fDist, sDist;
            int clusInd = 0;
            double[] fPoint = null, sPoint = null;
            foreach (ClusterInf cluster in dataList)
            {
                foreach (double[] firElem in cluster.Cluster)
                {
                    foreach (double[] secElem in cluster.Cluster)
                    {
                        if (firElem == secElem)
                            continue;
                        
                         
                        dist = 0;
                        for (int i = 0; i < firElem.Length; i++)
                        {
                            dist += Math.Pow((firElem[i] - secElem[i]), 2);
                        }
                        if (dist > maxDist)
                        {
                            maxDist = dist;
                            fPoint = firElem;
                            sPoint = secElem;
                            clusInd = dataList.IndexOf(cluster);
                        }
                    }
                }
            }
            ClusterInf nCluster = new ClusterInf();
            nCluster.Cluster = new List<double[]>();
            nCluster.Cluster.Add(sPoint);
            dataList[clusInd].Cluster.Remove(sPoint);
            foreach (double[] checkedElem in dataList[clusInd].Cluster)
            {
                if (checkedElem == fPoint)
                    continue;
                fDist = 0;
                sDist = 0;
                for (int i = 0; i < checkedElem.Length; i++)
                {
                    fDist += Math.Pow((checkedElem[i] - fPoint[i]), 2);
                    sDist += Math.Pow((checkedElem[i] - sPoint[i]), 2);
                }
                if (fDist > sDist)
                {
                    nCluster.Cluster.Add(checkedElem);
                    
                }
            }
            foreach (double[] point in nCluster.Cluster)
            {
                dataList[clusInd].Cluster.Remove(point);
            }
            dataList.Add(nCluster);
            return dataList;
        } 

        private static double[] FindInDeep(double[] x, List<double[]> VerifCluster, double minDist)
        {
            double dist = 0;
            bool changed = false;
            double [] newItem = new double[x.Length];
            foreach (double[] elOfClus in VerifCluster)
            {
                if(x == elOfClus)
                    continue;
                for (int i=0; i < x.Length; i++)
                {
                    dist += Math.Pow((x[i] - elOfClus[i]), 2);
                }
                if (dist < minDist)
                {
                    minDist = dist;
                    newItem = elOfClus;
                    changed = true;
                }
            }
            return newItem;
            //return null;
        }

        private List<ClusterInf> DoubleFindIerAlg(List<ClusterInf> UpIerData, List<ClusterInf> DownIerData)
        {
            double d = 0, minD = Double.MaxValue;
            List<ClusterInf> neededItUpCluster = new List<ClusterInf>();
            List<ClusterInf> neededItDownCluster = new List<ClusterInf>();

            while (UpIerData.Count != 2)
            {
                UpIerData = UpIerarchAlg(UpIerData);
                List<ClusterInf> downIerChData = new List<ClusterInf>();
                foreach (ClusterInf cluster in DownIerData)
                {
                     downIerChData.Add((ClusterInf)cluster.Clone());    
                }               

                while (UpIerData.Count != downIerChData.Count)
                {
                    downIerChData = DownIerarchAlg(downIerChData);
                }
                double sumK=0, subQ = 0, fpx = 0, fpy = 0; 
                foreach (ClusterInf upCluster in UpIerData)
                {
                    sumK = 0;
                    foreach (ClusterInf downCluster in downIerChData)
                    {
                        double samePointCount = 0;
                        sumK += Math.Pow(downCluster.Cluster.Count, 2);
                        foreach (double[] samePoints in downCluster.Cluster)
                        {
                            samePointCount += upCluster.HasElem(samePoints) ? 1 : 0;
                        }
                        fpy += Math.Pow(samePointCount, 2);
                        
                    }
                    subQ += Math.Pow(upCluster.Cluster.Count, 2);
                }
                fpx = (sumK + subQ)/2;
                d = (fpx - fpy)/fpx;
                if (d<minD)
                {
                    neededItDownCluster.Clear();
                    neededItUpCluster.Clear();
                    minD = d;
                    neededItUpCluster.AddRange(UpIerData.Select(cluster => (ClusterInf) cluster.Clone()));
                    neededItDownCluster.AddRange(downIerChData.Select(cluster => (ClusterInf) cluster.Clone())) ;
                }
               
            }
            if (neededItDownCluster == null || neededItUpCluster == null) return null;
            int[,] aMatrix = new int[neededItDownCluster.Count, neededItUpCluster.Count];
            List<ClusterInf> doubClustRes = new List<ClusterInf>();
            int col = 0, row = 0;
            foreach (ClusterInf Q in neededItUpCluster)
            {
                col = 0;
                ClusterInf cli = new ClusterInf();
                
                doubClustRes.Add(new ClusterInf());
                foreach (ClusterInf K in neededItDownCluster)
                {
                    foreach (double[] chekblElem in Q.Cluster)
                    {
                        if (K.HasElem(chekblElem))
                        {
                            ++aMatrix[row, col];
                            doubClustRes[row].Cluster.Add(chekblElem);
                        }
                    }
                    col++;
                }
                row++;
            }
            doubClustRes = clustOrg(doubClustRes, DownIerData);
            return doubClustRes;
        }

        private void UpIerarch_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            List<ClusterInf> doublClust = new List<ClusterInf>(DoubleFindIerAlg(upDataList, downDataList));
            UpIerarch.Series.Clear();
            int i = 0;
            foreach (ClusterInf cluster in doublClust)
            {
                UpIerarch.Series.Add(i.ToString());
                UpIerarch.Series[i].ChartType = SeriesChartType.Point;
                UpIerarch.Series[i].MarkerSize = 10;
                foreach (double[] point in cluster.Cluster)
                {
                    UpIerarch.Series[i].Points.AddXY(point[0], point[1]);
                }
                i++;
            }
        }

        private List<ClusterInf> clustOrg(List<ClusterInf> doublClustRes, List<ClusterInf> DownIerData)
        {
            double dist = 0, minDist = Double.MaxValue;
            
            foreach (double[] cheblPoint in DownIerData[0].Cluster)
            {
                int clusNum = -1;
                foreach (ClusterInf cluster in doublClustRes)
                {
                    if (cluster.HasElem(cheblPoint)) continue;
                    double[] midpoint = ClustMidPoint(cluster);
                    for (int i = 0; i < cheblPoint.Length; i++)
                    {
                        dist += Math.Pow((cheblPoint[i] - midpoint[i]), 2);
                    }
                    if (dist<minDist)
                    {
                        minDist = dist;
                        clusNum = doublClustRes.IndexOf(cluster);
                    }
                }
                if (clusNum == -1) continue;
                doublClustRes[clusNum].Cluster.Add(cheblPoint);
            }
            return doublClustRes;
        }

        private double[] ClustMidPoint(ClusterInf cluster)
        {
            double[] midPoint = new double[cluster.Cluster[0].Length];
            foreach (double[] point in cluster.Cluster)
            {
                for (int i = 0; i < point.Length; i++)
                {
                    midPoint[i] += point[i];
                }
            }
            for (int i = 0; i < midPoint.Length; i++)
            {
                midPoint[i] = midPoint[i]/cluster.Cluster.Count;
            }
            return midPoint;
        }
    }
}

