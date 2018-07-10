using System;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Web;

namespace Calculation.Models
{
    /// <summary>
    /// Http请求帮助类
    /// </summary>
    public static class HttpHelper
    {
        /// <summary>
        /// 缺省客户端代理
        /// </summary>
        private static readonly string DefaultUserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)";

        #region GET 请求
        /// <summary>
        /// 创建GET方式的HTTP请求，返回 Internet 资源响应 
        /// </summary>
        /// <param name="url">请求的URL</param>
        /// <param name="timeout">请求的超时时长（毫秒），为空则为系统默认最长时长</param>
        public static HttpWebResponse HttpGet(string url, int? timeout)
        {
            var request = CreateHttpRequest(url, "GET", null, null, timeout, null, null, null, null);
            return request.GetResponse() as HttpWebResponse;
        }
        #endregion


        #region POST 请求

        /// <summary>
        /// 创建POST方式的HTTP请求(采用UTF-8编码)，返回 Internet 资源响应
        /// </summary>
        /// <param name="url">请求的URL</param>
        /// <param name="parameters">随同请求POST的参数名称及参数值字典，可以为空</param>
        /// <param name="timeout">请求的超时时长（毫秒），为空则为系统默认最长时长</param>
        public static HttpWebResponse HttpPost(string url, NameValueCollection parameters, int? timeout)
        {
            return HttpPost(url, null, parameters, timeout, null, System.Text.Encoding.UTF8, null);
        }

        /// <summary>  
        /// 创建POST方式的HTTP请求，返回 Internet 资源响应
        /// </summary>  
        /// <param name="url">请求的URL</param>  
        /// <param name="headers">Http请求标头的键/值</param>
        /// <param name="parameters">随同请求POST的参数名称及参数值字典，可以为空</param>  
        /// <param name="timeout">请求的超时时长（毫秒），为空则为系统默认最长时长</param>  
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>  
        /// <param name="requestEncoding">发送HTTP请求时所用的编码</param>  
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param> 
        public static HttpWebResponse HttpPost(string url, NameValueCollection headers, NameValueCollection parameters, int? timeout, string userAgent, Encoding requestEncoding, CookieCollection cookies)
        {
            string contentType = "application/x-www-form-urlencoded";
            var request = CreateHttpRequest(url, "POST", headers, ToHttpParamsString(parameters, false), timeout, contentType, userAgent, requestEncoding, cookies);
            return request.GetResponse() as HttpWebResponse;
        }
        #endregion


        #region POST 请求（JSON）
        /// <summary>
        /// 创建POST方式的HTTP请求(采用UTF-8编码)，返回 Internet 资源响应
        /// </summary>
        /// <param name="url">请求的URL</param>
        /// <param name="json">JSON字符串（采用UTF-8编码） </param>
        /// <param name="timeout">请求的超时时长（毫秒），为空则为系统默认最长时长</param>
        public static HttpWebResponse HttpPostJson(string url, string json, int? timeout)
        {
            return HttpPostJson(url, null, json, timeout, null, null);
        }

        /// <summary>
        /// 创建POST方式的HTTP请求(采用UTF-8编码)，返回 Internet 资源响应  
        /// </summary>
        /// <param name="url">请求的URL</param>
        /// <param name="headers">Http请求标头的键/值</param>
        /// <param name="json">JSON字符串 (采用UTF-8编码) </param>
        /// <param name="timeout">请求的超时时长（毫秒），为空则为系统默认最长时长</param>
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>
        public static HttpWebResponse HttpPostJson(string url, NameValueCollection headers, string json, int? timeout, string userAgent, CookieCollection cookies)
        {
            string contentType = "application/json;charset=UTF-8";
            var request = CreateHttpRequest(url, "POST", headers, json, timeout, contentType, userAgent, System.Text.Encoding.UTF8, cookies);
            return request.GetResponse() as HttpWebResponse;
        }
        #endregion

        /// <summary>
        /// 创建POST方式的HTTP请求
        /// </summary>
        /// <param name="url">请求的URL</param>
        /// <param name="method">HTTP请求 Method</param>
        /// <param name="headers">Http请求标头的键/值</param>
        /// <param name="dataStr">请求的参数</param>
        /// <param name="timeout">请求的超时时长（毫秒），为空则为系统默认最长时长</param>
        /// <param name="contentType">HTTP请求 Content-Type</param>
        /// <param name="userAgent">请求的客户端浏览器信息，可以为空</param>
        /// <param name="requestEncoding">发送HTTP请求时所用的编码，可以为空（默认UTF-8）</param>
        /// <param name="cookies">随同HTTP请求发送的Cookie信息，如果不需要身份验证可以为空</param>
        private static HttpWebRequest CreateHttpRequest(string url, string method, NameValueCollection headers, 
                                                        string dataStr, int? timeout, string contentType,  
                                                        string userAgent, Encoding requestEncoding, CookieCollection cookies)
        {
            #region 必要参数校验
            if (string.IsNullOrEmpty(url))
            {
                throw new ArgumentNullException("url");
            }
            #endregion

            #region 针对HTTPS请求进行处理
            HttpWebRequest request = null;
            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                //如果是发送HTTPS请求  
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version10;
            }
            else
            {
                request = WebRequest.Create(url) as HttpWebRequest;
            }
            #endregion

            #region 设置 userAgent、timeout、cookies、header、method、contentType
            request.UserAgent = !string.IsNullOrEmpty(userAgent) ? userAgent : DefaultUserAgent;

            if (timeout.HasValue)
            {
                request.Timeout = timeout.Value;
            }
            if (cookies != null)
            {
                request.CookieContainer = new CookieContainer();
                request.CookieContainer.Add(cookies);
            }

            if (headers != null)
            {
                request.Headers.Add(headers);
            }

            if (!string.IsNullOrEmpty(contentType))
            {
                request.ContentType = contentType;
            }

            request.Method = method;
            #endregion


            if (!string.IsNullOrEmpty(dataStr))
            {
                byte[] data = requestEncoding == null ? System.Text.Encoding.UTF8.GetBytes(dataStr) : requestEncoding.GetBytes(dataStr);
                request.ContentLength = data.Length;
                using (Stream stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            return request;
        }

        #region 私有方法
        /// <summary>
        /// 验证证书
        /// </summary>
        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            if (errors == SslPolicyErrors.None)
                return true;
            return false;
        }
        private static string ToHttpParamsString(NameValueCollection parameters, bool urlencoded)
        {
            if (parameters == null || parameters.Count == 0)
            {
                return " ";// string.Empty;
            }
            StringBuilder builder = new StringBuilder();

            foreach (string key in parameters.Keys)
            {
                if (builder.Length > 0)
                {
                    builder.Append('&');
                }

                string keyStr = (urlencoded ? HttpUtility.UrlEncode(key) : key) + "=";

                string[] values = parameters.GetValues(key);
                int valueCounts = values != null ? values.Length : 0;
                if (valueCounts == 0)
                {
                    builder.Append(keyStr);
                }
                else if (valueCounts == 1)
                {
                    builder.Append(keyStr);
                    string valueStr = urlencoded ? HttpUtility.UrlEncode(values[0]) : values[0];
                    builder.Append(valueStr);
                }
                else
                {
                    for (int i = 0; i < valueCounts; i++)
                    {
                        if (i > 0)
                        {
                            builder.Append('&');
                        }
                        builder.Append(keyStr);
                        string valueStr = urlencoded ? HttpUtility.UrlEncode(values[i]) : values[i];
                        builder.Append(values[i]);
                    }
                }
            }
            return builder.ToString();
        }
        #endregion


        /// <summary>
        /// 获取请求的数据
        /// </summary>
        public static string GetResponseString(HttpWebResponse webresponse)
        {
            using (Stream s = webresponse.GetResponseStream())
            {
                StreamReader reader = new StreamReader(s, Encoding.UTF8);
                return reader.ReadToEnd();
            }
        }
    }
}
