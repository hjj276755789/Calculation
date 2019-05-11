using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace Calculation
{
    public class SocketStore
    {
        public string key { get; set; }
        public Socket sc { get; set; }

        public DateTime time { get; set; }
        public int gqbz { get; set; }
    }
    /// <summary>
    /// WebSocket线程池
    /// </summary>
    public class SocketStoreManager
    {
        private static Thread refth;  //刷新线程

        public void Initialization()
        {
            refth = new Thread(Refresh);
            refth.IsBackground = true;
            refth.Start();
        }

        /// <summary>
        /// 线程池操作锁
        /// </summary>
        public static object s_lock = new object();
        /// <summary>
        /// WebSocket线程池实体
        /// </summary>
        public static List<SocketStore> pool { get; set; }

        /// <summary>
        /// 添加
        /// </summary>
        public static void Add(SocketStore entity)
        {
            lock (s_lock)
            {
                if (pool == null)
                    pool = new List<SocketStore>();
                //如果池中没有对应的实体，则添加，否则跳过
                if (!pool.Exists(m => m.key == entity.key))
                    pool.Add(entity);
                else
                    Console.WriteLine("重复实体");

            }

        }
        /// <summary>
        /// 删除
        /// </summary>
        /// <param name="entity"></param>
        public static void Del(SocketStore entity)
        {
            lock (s_lock)
            {
                try
                {
                    if (pool != null)
                        pool.Remove(entity);
                }
                catch (Exception)
                {

                    Console.WriteLine("删除池子报错");
                }
            }
        }

        public static SocketStore Find(string key)
        {
            lock (s_lock)
            {
                if (pool != null)
                    return pool.FirstOrDefault(m => m.key == key);
                else return null;
            }
        }

        /// <summary>刷新</summary>
        /// 锁机制有问题，需要调整
        public void Refresh()
        {
            lock (s_lock)
            {
                Console.WriteLine("开始刷新");
                if (pool != null)
                {
                    List<SocketStore> temp = new List<SocketStore>(); ;
                    foreach (var item in pool)
                    {
                        int s = (int)(DateTime.Now - item.time).TotalSeconds;
                        Console.WriteLine(item.key + ":" + s);
                        if (s > 40)
                        {
                            temp.Add(item);
                        }
                    }
                    if (temp != null)
                    {
                        foreach (var item in temp.ToArray())
                        {
                            Console.WriteLine("超时连接：" + item.key + "：当前时常" + (int)(DateTime.Now - item.time).TotalSeconds);
                            item.gqbz = 1;
                            pool.Remove(item);
                        }
                    }
                    temp = null;
                }
            }
        }

    }
}
