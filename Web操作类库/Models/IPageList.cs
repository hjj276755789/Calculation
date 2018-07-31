using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models
{
    public interface IPageList<T> : IList<T>
    {
        /// <summary>
        /// 获取或设置页码。
        /// </summary>
        int PageNumber { get; set; }
        /// <summary>
        /// 获取或设置数据项大小。
        /// </summary>
        int PageSize { get; set; }
        /// <summary>
        /// 获取或设置数据项总数。
        /// </summary>
        int TotalItemCount { get; set; }
        /// <summary>
        /// 获取数据项总页数。
        /// </summary>
        int TotalPageCount { get; }
        /// <summary>
        /// 获取起始位置。
        /// </summary>
        int StartPosition { get; }
        /// <summary>
        /// 获取结束位置。
        /// </summary>
        int EndPosition { get; }
        /// <summary>
        /// 获取一个值，该值表示是否还有页面存在于当前页之前。
        /// </summary>
        bool HasPreviousPage { get; }
        /// <summary>
        /// 获取一个值，该值表示是否还有页面存在于当前页之后。
        /// </summary>
        bool HasNextPage { get; }

    }
    public class PageList<T> : List<T>, IPageList<T>
    {
        public PageList(IEnumerable<T> items, int pageNumber, int pageSize, int totalItemCount)
        {
            AddRange(items);
            this.PageNumber = pageNumber;
            this.PageSize = pageSize;
            this.TotalItemCount = totalItemCount;
        }

        internal PageList(IEnumerable<T> items)
        {
            AddRange(items);
        }
        /// <summary>
        /// 获取或设置页码。
        /// </summary>
        public int PageNumber { get; set; }
        /// <summary>
        /// 获取或设置数据项大小。
        /// </summary>
        public int PageSize { get; set; }
        /// <summary>
        /// 获取或设置数据项总数。
        /// </summary>
        public int TotalItemCount { get; set; }
        /// <summary>
        /// 获取数据项总页数。
        /// </summary>
        public int TotalPageCount
        {
            get { return (int)Math.Ceiling((double)TotalItemCount / PageSize); }
        }
        /// <summary>
        /// 获取起始位置。
        /// </summary>
        public int StartPosition
        {
            get { return (PageNumber - 1) * PageSize + 1; }
        }
        /// <summary>
        /// 获取结束位置。
        /// </summary>
        public int EndPosition
        {
            get { return PageNumber * PageSize > TotalItemCount ? TotalItemCount : PageNumber * PageSize; }
        }
        /// <summary>
        /// 获取一个值，该值表示是否还有页面存在于当前页之前。
        /// </summary>
        public bool HasPreviousPage
        {
            get { return (PageNumber > 1); }
        }
        /// <summary>
        /// 获取一个值，该值表示是否还有页面存在于当前页之后。
        /// </summary>
        public bool HasNextPage
        {
            get { return (PageNumber * PageSize) < TotalItemCount; }
        }
    }

}
