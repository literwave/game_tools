# game_tools
游戏开发里面的工具：例如：导表工具

### 导表工具

Export excel to Python/Json/Lua

参考 `https://github.com/hanxi/export.py.git`

* 使用方法

  * 运行`导表脚本.bat`便可以客户端和服务端生成所需要的文件

* 目录层级

  * `client` 客户端所需要的配置文件
  * `server`服务端所需要的配置文件
  * `excel`配置表

* 生成格式

  ```
  第一行第一列 代表生成配置文件的名字 例如hero.xlsx中的hero表示生成hero.xxx
  第二行表示注释
  第三行表示类型(支持table)
  第四行表示前后端所需要的 字段名
  第五行表示生成前端还是后端所需 `c/s`表示前后端都需要 
  ```

* `python export_file.py -r ./excel/hero.xlsx -f lua -t ../server -o s` 可以手动调试，里面有调式打印代码，可查看

* `hero.xlsx`导出的是这样的

  ```lua
  -- author:   liter_wave 
  -- Automatic generation from -->>
  -- excel file  name: ./excel/hero.xlsx
  -- excel sheet name: 英雄
  return {
      [2] = {
          ["Name"] = "奥丁",
          ["MountId"] = 10001,
          ["Sex"] = 1,
          ["num"] = {
              1,
              2,
              3
          },
          ["zzz"] = {
              1,
              2,
              3
          },
          ["ddd"] = {{1,2,3},{2,3,4}},
          ["xxx"] = {1,2,3}
      }
  }
  
  ```

* 第三行注释表示文件目录，第四行表示哪张子表，可定位到出错的位置。

* 导表工具还有许多优化的地方，比如；在`bat`脚本中可定位到错误
