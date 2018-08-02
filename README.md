# # IP Block Design

在日常工作中，经常需要划分`IP`网段，一般客户会给我们提供一个大段的私网地址，比如`18`位掩码`IP`，如果使用手工分配的话，会耗费大量的精力，也容易出错，因此写了这个小工具，可以自动完成`IP`设计。

## 依赖库

- IPy: `pip3 install IPy`
- xlsxwriter: `pip3 install xlsxwriter`



## 使用步骤

在使用中只需要根据实际情况修改程序的三个参数。

- max_mask : 获取的掩码位数，如18位掩码
- min_mask :  自己需要规划的最小单位的掩码，比如30位掩码
- original_ip_block :初始化IP对象， IP('10.75.96.0')

然后执行程序即可在当前目录生成**IP_blocks.xlxl**文件：

```shell
$python3 ip_block_design.py
```

![示意图](https://github.com/wowmarcomei/ip_block_design/snapshot.png)