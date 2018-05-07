using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Experiment
{
    public enum OutputType
    {
        All,//引用+自引
        Self,//引用+自引
        Other,//引用+他引
        Other2,//只有他引
        SelfWithMatchedAuthors,//包含作者匹配信息
        Test//测试
    }
}
