# 更新任务栏图标进度

## 安装模块
```sh
python -m pip install pywin32
pip install install https://github.com/vasily-v-ryabov/comtypes/archive/refs/heads/mtime_none.zip
```
## 解决报错
1. **`ImportError: Typelib different than module`**

解决方法：
```sh
pip uninstall -y comtypes
pip install install https://github.com/vasily-v-ryabov/comtypes/archive/refs/heads/mtime_none.zip
# 清除缓存
python /python安装目录/Scripts/clear_comtypes_cache.py
```
运行再次报错：`ModuleNotFoundError: No module named 'comtypes.gen.TaskbarLib'`
解决办法：

在脚本引用 `TaskbarLib` 前，先运行如下代码，它会在 `Lib\site-packages\comtypes\gen` 下生成我们要用的这个 `TaskbarLib`
```
from comtypes.client import GetModule
GetModule(os.path.join(sys.path[0], 'lib/tl.tlb'))
```
之后运行还会报错一次，再次运行发现不报错了。

## 参考：
- [ImportError: Typelib different than module · Issue #231 · enthought/comtypes](https://github.com/enthought/comtypes/issues/231#issuecomment-841767788)
