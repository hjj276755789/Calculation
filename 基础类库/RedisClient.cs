using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StackExchange.Redis;

namespace Calculation.Base
{
    public static class RedisClient
    {
        private static ConnectionMultiplexer connection;
        private static ConnectionMultiplexer _Connection
        {
            get
            {
                if (connection == null || !connection.IsConnected)
                {
                    connection = ConnectionMultiplexer.Connect(_ConnectionOptions);
                    redis = _Connection.GetDatabase();
                }
                return connection;
            }
            set
            {
                connection = value;
            }
        }
        private static ConfigurationOptions connectionOptions;

        private static ConfigurationOptions _ConnectionOptions
        {
            get
            {
                //if (connectionOptions == null) connectionOptions = ConfigurationOptions.Parse(System.Configuration.ConfigurationManager.ConnectionStrings["redis"].ConnectionString);
                if (connectionOptions == null) connectionOptions = ConfigurationOptions.Parse(System.Configuration.ConfigurationManager.ConnectionStrings["redis"].ConnectionString);
                return connectionOptions;
            }
            set
            {
                connectionOptions = value;
            }
        }
        public static IDatabase redis;
        private static IDatabase _Redis
        {
            get
            {
                if (redis == null) redis = _Connection.GetDatabase();
                return redis;
            }
            set
            {
                redis = value;
            }
        }

        static RedisClient()
        {
            _ConnectionOptions = ConfigurationOptions.Parse(System.Configuration.ConfigurationManager.ConnectionStrings["redis"].ConnectionString);
            _Connection = ConnectionMultiplexer.Connect(_ConnectionOptions);
            _Connection.PreserveAsyncOrder = true;
            _Redis = _Connection.GetDatabase();
        }

        public static bool KeyExists(string name)
        {
            return _Redis.KeyExists(name);
        }


        //"__keyevent@0__:expired"

        public static void SetSubscribe(string channelName, Action<string, string> handler)
        {
            _Connection.GetSubscriber()
                .Subscribe(
                new RedisChannel(channelName, RedisChannel.PatternMode.Auto),
                (channel, value) =>
                {
                    handler(channel, value);
                }
                );
        }

        public static void KeyExpire(string name, DateTime? expire)
        {
            _Redis.KeyExpire(name, expire);
        }

        public static RedisValue[] Set(string name)
        {
            return _Redis.SetMembers(name);
        }

        public static bool Set(string name, RedisValue data, DateTime? expire = null)
        {
            _Redis.SetAdd(name, data);
            return _Redis.KeyExpire(name, expire);
        }

        public static long Set(string name, RedisValue[] data, DateTime? expire = null)
        {
            _Redis.KeyExpire(name, expire);
            return _Redis.SetAdd(name, data);
        }

        public static void String(string name, RedisValue data, DateTime? expire = null)
        {
            _Redis.StringSet(name, data);
            _Redis.KeyExpire(name, expire);
        }

        public static string String(string name)
        {
            return _Redis.StringGet(name);
        }

        public static string StringGetSet(string name, RedisValue value)
        {
            return _Redis.StringGetSet(name, value);
        }

        public static bool StringExists(string name, RedisValue value)
        {
            return _Redis.StringGet(name) == value;
        }

        public static void Hash(string name, HashEntry[] data, DateTime? expire = null)
        {

            _Redis.HashSet(name, data);
            _Redis.KeyExpire(name, expire);
        }

        public static string HashGet(string name, RedisValue field)
        {
            return _Redis.HashGet(name, field);
        }

        public static Dictionary<string, string> HashGet(string name)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            var values = _Redis.HashGetAll(name);
            foreach (var item in values)
            {
                result.Add(item.Name, item.Value);
            }
            return result;
        }

        public static bool Hash(string name, RedisValue field, RedisValue value, DateTime? expire = null)
        {

            _Redis.HashSet(name, field, value);
            return _Redis.KeyExpire(name, expire);
        }

        public static RedisValue[] HashGetArray(string name, RedisValue[] fields)
        {
            return _Redis.HashGet(name, fields);
        }

        public static bool HashDelete(string name, RedisValue field)
        {
            return _Redis.HashDelete(name, field);
        }
    }
}
