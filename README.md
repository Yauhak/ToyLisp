# ToyLisp - 一个轻量级Lisp方言解释器

**ToyLisp** 是一个用VB6实现的类Lisp语法脚本语言解释器，诞生于作者高二时期的编程探索。它支持函数式编程、基础数据结构操作，甚至能实现简单的神经网络！虽名为“玩具”，但功能齐全，适合学习或小型项目。

## 📦 功能特性
- **Lisp风格语法**：`(操作符 参数...)` 的简洁表达式
- **丰富的基础功能**：
  - 数学运算（`+ - * / sqrt sin`等）
  - 条件分支（`if`）、循环（`while`）
  - 字符串操作（`&`拼接、`substr`截取）
  - 数组操作（`list`创建、`m`索引、`array`修改）
  - 文件读写（`read`/`outfile`）
  - 自定义函数（`fn`）与公有变量（`public`）
- **跨层级示例**：从斐波那契数列到BP神经网络
- **错误提示**：友好的运行时错误诊断

## 🚀 快速开始
### 运行环境
- Windows系统 + VB环境

### 基础示例
```lisp
(main (out "Hello ToyLisp!"))
```
```lisp
(# 阶乘)
(fn factorial (n)
(if (<= n 1)
     (return  1)
     (return (* n (factorial (- n 1))))
)
(main (out (factorial 5)))(# 输出120)
```
### 📚 语法
- 变量定义
```lisp
(def pi 3.1415)  (# 定义变量)
(public
    MAX_SIZE 100  (# 全局变量)
)
```
- 数组操作
```lisp
(def l(list 1 2 3 (list 4 5))) (# 定义一个数组、列表)
(out (m l 3 0)) (# 等义于l[3][0]，此时值为4)
(array l (2) 6) (# 将l[2]的值改为6，注意此处下标要用括号括起来，如果是要访问的是多维下标则写成(x1 x2 ...))
```
- 神经网络示例-三层神经网络（见BP神经网络-三层.lsp）
```lisp
(main
   (def syn0 (list (rand 1) (rand 1)...) (# 初始化权重)
   (while (< iter 1000)  (# 迭代训练)
      (def l1 (active (matrix_mul m1 syn0) 1)
      ... (# 反向传播逻辑)
   )
)
```
### 🎯 示例程序
#### 示例名称	              功能描述
##### 斐波那契数列.lsp	    生成斐波那契序列
##### 九九乘法口诀表.lsp	  输出格式化乘法表
##### 最小二乘法-高级版.lsp	从文件读取数据拟合线性方程
##### BP神经网络-三层.lsp	  实现异或运算的三层神经网络

### 🛠️ 使用指南
- 编写脚本：用.lsp后缀保存文件
- 运行方式：关联.lsp文件到ToyLisp.exe
- （可以直接将编写好的ToyLisp文件拖动到编译器上运行啦）

### ⚠️ 注意事项
- 函数/变量作用域：同名定义以首次出现为准
- 错误处理：目前仍较简单，复杂运算建议添加校验
- 性能限制：迭代超过1000次可能较慢

### 🤝 参与贡献
- 欢迎提交Issue或PR！建议方向：
- 优化错误处理机制
- 增加更多数学库函数
- 提升解释器执行效率
(用其他编程语言重写也不是不行(lll￢ω￢)）

## 📮 联系作者
### QQ: 3953814837
### 华子（Yauhak）
