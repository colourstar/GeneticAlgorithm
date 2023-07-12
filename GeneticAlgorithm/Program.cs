﻿using ExcelGroupCalculater;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace GeneticAlgorithm
{
    public class Program
    {
        public class Position
        {
            public float m_fPositionX = 0.0f;
            public float m_fPositionY = 0.0f;

            public float Distance(Position pos)
            {
                return (m_fPositionX - pos.m_fPositionX) * (m_fPositionX - pos.m_fPositionX) +
                    (m_fPositionY - pos.m_fPositionY) * (m_fPositionY - pos.m_fPositionY);
            }
        }
        // 定义分组后的每个楼号
        public class Building
        {
            public string m_strName = "";       // 楼号名
            public Position m_vecPosition = new Position();
            public float m_fWValue = 0.0f;      // 负荷
        }

        // 定义个体
        public class Individural
        {
            public List<Position> m_kPositions = new List<Position>();

            public float m_fFitness = 0.0f;

            public void Clone(Individural origin)
            {
                m_kPositions.Clear();
                m_fFitness = origin.m_fFitness;
                origin.m_kPositions.ForEach(t => 
                {
                    m_kPositions.Add(new Position() { m_fPositionX = t.m_fPositionX, m_fPositionY = t.m_fPositionY });
                });
            }
        }

        public static List<Building> m_arrBuildingList = new List<Building>();
        public static List<Individural> m_arrIndividural = new List<Individural>();
        private static readonly int m_iIndividuarlNumbers = 50;         // 个体数量
        private static readonly int m_iSelectNumbers = 50;              // 每轮生成新个体的数量
        private static readonly int m_iGenerateCount = 1000;             // 迭代次数
        private static readonly float m_fMaxPositionX = 412;            // x最大范围
        private static readonly float m_fMaxPositionY = 258;            // y最大范围
        private static readonly int m_iElectricNodeNumbers = 8;         // 变电所数量

        private static readonly float CiSi = 150.0f;                    // Ci(Si)
        private static readonly float R0 = 8;
        private static readonly int m = 20;
        private static readonly float CuSi = 10.0f;
        private static readonly float alpha = 0.07f;
        private static readonly int k = 20;
        private static readonly float beta = 0.12f;
        private static readonly float cosinBeta = 0.92f;

        private static readonly float m_fR0MParam = (R0 * (float)Math.Pow(1.0f + R0, m)) / (R0 * (float)Math.Pow(1.0f + R0, m) - 1.0f);
        private static readonly float m_fR0KParam = (R0 * (float)Math.Pow(1.0f + R0, k)) / (R0 * (float)Math.Pow(1.0f + R0, k) - 1.0f);



        private static Random m_kRandom = new Random();
        public static Dictionary<int, List<int>> m_dicIndividuralToBuilding = new Dictionary<int, List<int>>();

        public static void Main(string[] args)
        {
            // 初始化所有楼
            _InitBuilding();

            // 初始化个体
            _InitIndividural();

            int iIndex = 0;
            while (iIndex < m_iGenerateCount)
            {
                // 针对每一个个体计算其适应度
                _CaculateFitness();

                // 选择
                _SelectIndividural();

                // 交叉
                _CrossIndividural();

                // 变异
                _VariationIndividural();

                // 禁忌算法
                if (iIndex % 10 == 0)
                {

                }

                // 对适应度再次进行排序
                _SortByFitness();

                // 输出结果
                _OutputResult(iIndex);


                iIndex++;
            }

            // 停留等待
            Console.Read();
        }
        #region InitFrame
        /// <summary>
        /// 从Excel中读取楼号和位置信息
        /// </summary>
        private static void _InitBuilding()
        {
            m_arrBuildingList.Clear();
            DataRowCollection kDataCollection = ExcelHelper.ReadExcel($"{System.Environment.CurrentDirectory}\\Excel.xlsx", "Sheet1");
            for (int i = 0; i < kDataCollection.Count; ++i)
            {
                string strName = kDataCollection[i][0].ToString();
                string strValue = kDataCollection[i][1].ToString();
                string strPositionX = kDataCollection[i][2].ToString();
                string strPositionY = kDataCollection[i][3].ToString();

                if (float.TryParse(strValue, out float iValue) == false)
                {
                    continue;
                }

                if (float.TryParse(strPositionX, out float fPositionX) == false ||
                    float.TryParse(strPositionY, out float fPositionY) == false)
                {
                    continue;
                }

                Building kCell = new Building()
                {
                    m_strName = strName,
                    m_vecPosition = new Position() { m_fPositionX = fPositionX, m_fPositionY = fPositionY },
                    m_fWValue = iValue
                };
                m_arrBuildingList.Add(kCell);
            }

            // 位置对应楼号的方案是确定的
            m_dicIndividuralToBuilding.Add(0, new List<int>() { 1, 8, 9, 18 });
            m_dicIndividuralToBuilding.Add(1, new List<int>() { 2, 3, 4 });
            m_dicIndividuralToBuilding.Add(2, new List<int>() { 5, 12 });
            m_dicIndividuralToBuilding.Add(3, new List<int>() { 6, 7, 11 });
            m_dicIndividuralToBuilding.Add(4, new List<int>() { 10, 17 });
            m_dicIndividuralToBuilding.Add(5, new List<int>() { 13,14, 21 });
            m_dicIndividuralToBuilding.Add(6, new List<int>() { 15, 20 });
            m_dicIndividuralToBuilding.Add(7, new List<int>() { 16, 19 });
        }

        /// <summary>
        /// 初始化遗传算法个体
        /// </summary>
        private static void _InitIndividural()
        {
            // 创建100个个体,每个个体里拥有8个变电所的随机位置
            for (int i = 0; i < m_iIndividuarlNumbers; ++i)
            {
                var instance = new Individural();
                for (int iIndex = 0; iIndex < m_iElectricNodeNumbers; ++iIndex)
                {
                    instance.m_kPositions.Add(GenerateRandomPos());
                }
                m_arrIndividural.Add(instance);
            }
        }

        /// <summary>
        /// 范围内随机一个位置
        /// </summary>
        /// <returns></returns>
        private static Position GenerateRandomPos()
        {
            float fPositionX = (float)m_kRandom.Next(0, (int)m_fMaxPositionX);
            float fPositionY = (float)m_kRandom.Next(0, (int)m_fMaxPositionY);

            return new Position() { m_fPositionX = fPositionX, m_fPositionY = fPositionY};
        }
        #endregion

        #region Fitness
        /// <summary>
        /// 计算所有个体的适应度
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        private static void _CaculateFitness()
        {
            for (int i = 0; i < m_arrIndividural.Count; ++i)
            {
                m_arrIndividural[i].m_fFitness = _CaculateIndividural(m_arrIndividural[i]);
            }
        }

        /// <summary>
        /// 计算某一个个体的适应度
        /// </summary>
        /// <param name="instance"></param>
        /// <returns></returns>
        private static float _CaculateIndividural(Individural instance)
        {
            float fFitness = 0.0f;
            // 适应度的算法，暂时采用所在变电所到达各个供电单元的距离


            // 计算方程式部分一
            float fFitnessPart1 = 0.0f;
            for (int i = 0; i < instance.m_kPositions.Count; ++i)
            {
                float fFitnessSingle = CiSi * m_fR0MParam + CuSi;
                fFitnessPart1 += fFitnessSingle;
            }

            // 计算方程式部分二
            float fFitnessPart2 = 0.0f;
            for (int i = 0; i < instance.m_kPositions.Count; ++i)
            {
                // 每一个距离都乘一下
                for (int iBuildingIndex = 0; iBuildingIndex < m_iElectricNodeNumbers; ++iBuildingIndex)
                {
                    Position kElectricPos = instance.m_kPositions[iBuildingIndex];
                    m_dicIndividuralToBuilding[i].ForEach(t => 
                    {
                        Position kBuildingPos = m_arrBuildingList[t - 1].m_vecPosition;

                        float fDistance = kElectricPos.Distance(kBuildingPos);
                        fFitnessPart2 += (float)Math.Sqrt(fDistance);
                    });
                }
            }
            fFitnessPart2 = fFitnessPart2 * alpha * m_fR0KParam;

            // 计算方程式部分三
            float fFitnessPart3 = 0.0f;
            for (int i = 0; i < instance.m_kPositions.Count; ++i)
            {
                // 每一个距离都乘一下, 最后再乘以一下负荷
                for (int iBuildingIndex = 0; iBuildingIndex < m_iElectricNodeNumbers; ++iBuildingIndex)
                {
                    Position kElectricPos = instance.m_kPositions[iBuildingIndex];

                    m_dicIndividuralToBuilding[i].ForEach(t =>
                    {
                        Position kBuildingPos = m_arrBuildingList[t - 1].m_vecPosition;

                        float fDistance = kElectricPos.Distance(kBuildingPos);
                        fFitnessPart3 += (float)Math.Sqrt(fDistance) * m_arrBuildingList[t - 1].m_fWValue;
                    });
                }
            }
            fFitness = fFitnessPart1 + fFitnessPart2 + fFitnessPart3;

            return fFitness;
        }

        private static void _SortByFitness()
        {
            m_arrIndividural.Sort((a, b) => 
            {
                return a.m_fFitness.CompareTo(b.m_fFitness);
            });
            m_arrIndividural.Reverse();
        }
        #endregion

        #region SelectParent
        /// <summary>
        /// 利用轮盘赌算法,计算出来更可能的
        /// </summary>
        private static void _SelectIndividural()
        {
            // 首先对所有的数组进行一下从大到小的排序
            m_arrIndividural.Sort((a, b) =>
            {
                return a.m_fFitness.CompareTo(b.m_fFitness);
            });
            m_arrIndividural.Reverse();

            // 计算总的适应度
            float fTotalFitness = 0.0f;
            m_arrIndividural.ForEach(t => { fTotalFitness += t.m_fFitness; });

            // 孩子节点
            List<Individural> arrChildIndividural = new List<Individural>();

            // 进行迭代计算,
            for (int i = 0; i < m_iIndividuarlNumbers; ++i)
            {
                float fRandomFitness = m_kRandom.Next(0, (int)fTotalFitness);
                int iIndex = _NextRouletteIndex(fRandomFitness, m_arrIndividural);

                if (iIndex < 0 || iIndex >= m_arrIndividural.Count)
                {
                    continue;
                }

                Individural childInstance = new Individural();
                childInstance.Clone(m_arrIndividural[iIndex]);
                arrChildIndividural.Add(childInstance);
            }

            // 迭代完毕之后,生成新的一代,直接赋值过去
            m_arrIndividural.Clear();
            arrChildIndividural.ForEach(t => { m_arrIndividural.Add(t); });
        }


        /// <summary>
        /// 轮盘赌算法,选中元素的索引
        /// </summary>
        /// <param name="fTotalFitness"></param>
        /// <param name="arrCells"></param>
        /// <returns></returns>
        private static int _NextRouletteIndex(float fRandomFitness, List<Individural> arrCells)
        {
            for (int i = 0; i < arrCells.Count; ++i)
            {
                if (fRandomFitness < arrCells[i].m_fFitness)
                {
                    return i;
                }
                fRandomFitness -= arrCells[i].m_fFitness;
            }

            return -1;
        }
        #endregion

        #region Cross
        /// <summary>
        /// 交叉操作
        /// </summary>
        private static void _CrossIndividural()
        {
            // 交叉方式采用算术交叉
            for (int i = 0; i < m_arrIndividural.Count; i += 2)
            {
                // 重新计算交叉率
                float fPercent = _CaculateCurrentCrossPercent(i / 2, m_arrIndividural.Count / 2);

                if (fPercent < (float)m_kRandom.NextDouble())
                {
                    continue;
                }

                // 这里进行交叉操作,暂时做成8个变电所全部进行交叉
                for (int iIndex = 0; iIndex < m_iElectricNodeNumbers; ++iIndex) 
                {
                    float fPositionAX = m_arrIndividural[i].m_kPositions[iIndex].m_fPositionX;
                    float fPositionAY = m_arrIndividural[i].m_kPositions[iIndex].m_fPositionY;

                    float fPositionBX = m_arrIndividural[i + 1].m_kPositions[iIndex].m_fPositionX;
                    float fPositionBY = m_arrIndividural[i + 1].m_kPositions[iIndex].m_fPositionY;

                    float fPositionAXNew = fPercent * fPositionAX + (1.0f - fPercent) * fPositionBX;
                    float fPositionAYNew = fPercent * fPositionAY + (1.0f - fPercent) * fPositionBY;

                    float fPositionBXNew = fPercent * fPositionBX + (1.0f - fPercent) * fPositionAX;
                    float fPositionBYNew = fPercent * fPositionBY + (1.0f - fPercent) * fPositionAY;

                    m_arrIndividural[i].m_kPositions[iIndex].m_fPositionX = fPositionAXNew;
                    m_arrIndividural[i].m_kPositions[iIndex].m_fPositionY = fPositionAYNew;

                    m_arrIndividural[i + 1].m_kPositions[iIndex].m_fPositionX = fPositionBXNew;
                    m_arrIndividural[i + 1].m_kPositions[iIndex].m_fPositionY = fPositionBYNew;
                }
            }
        }

        /// <summary>
        /// 计算交叉率
        /// </summary>
        /// <param name="fCurrentPercent">当前交叉率</param>
        /// <param name="iCurrentCount">当前迭代次数</param>
        /// <param name="iTotalCount">当前总次数</param>
        /// <param name="fMinPercent">最小交叉率</param>
        /// <param name="fMaxPercent">最大交叉率</param>
        /// <returns></returns>
        private static float _CaculateCurrentCrossPercent(int iCurrentCount, int iTotalCount, float fMinPercent = 0.5f, float fMaxPercent = 1.0f)
        {
            float fCountPercent = (float)iCurrentCount / (float)iTotalCount;
            float fPercent = fMaxPercent * (float)Math.Cos(Math.PI / 2 * fCountPercent);

            if (fPercent == 1.0f)
            {
                fPercent = fMinPercent;
            }

            return fPercent;
        }
        #endregion

        #region Variation
        /// <summary>
        /// 变异操作
        /// </summary>
        private static void _VariationIndividural()
        {

        }
        #endregion

        #region Others
        /// <summary>
        /// 反射深度拷贝
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T DeepCopyReflection<T>(T obj)
        {
            var type = obj.GetType();
            object o = Activator.CreateInstance(type);
            System.Reflection.PropertyInfo[] PI = type.GetProperties();
            for (int i = 0; i < PI.Count(); i++)
            {
                System.Reflection.PropertyInfo P = PI[i];
                P.SetValue(o, P.GetValue(obj));
            }
            return (T)o;
        }

        private static void _OutputResult(int iTime)
        {
            Console.WriteLine($"第{iTime + 1}次迭代，最佳适应度为: {(int)m_arrIndividural[0].m_fFitness}, 变电所的位置为：");
            for (int i = 0; i < m_iElectricNodeNumbers; ++i)
            {
                Console.WriteLine($"变电所 {i + 1}，位置X：{m_arrIndividural[0].m_kPositions[i].m_fPositionX} 位置Y：{m_arrIndividural[0].m_kPositions[i].m_fPositionY}");
            }
        }
        #endregion
    }
}