(# 带了文件读入测试)
(# params.txt即坐标参数文件)
(fn fcal (lst idx)
   (def i 0)(def avx 0)
   (while (<= i idx)(
      (def avx (+ avx (m lst i 0)))
      (def i (+ i 1)))
   )(# X坐标之和)
   (def i 0)(def avy 0)
   (while (<= i idx)(
      (def avy (+ avy (m lst i 1)))
      (def i (+ i 1)))
   )(# Y坐标之和)
   (def avx (/ avx (+ idx 1)))
   (def avy (/ avy (+ idx 1)))
   (def j 0)(def sum_c 0)
   (while (<= j idx)(
      (def sum_c (+ sum_c 
      (*(- (m lst j 0) avx)
      (- (m lst j 1) avy))))
      (def j (+ j 1)))
   )
   (def j 0)(def sum_f 0)
   (while (<= j idx)(
      (def sum_f (+ sum_f 
      (*(- (m lst j 0) avx)
      (- (m lst j 0) avx))))
      (def j (+ j 1)))
   )
   (def _k_ (/ sum_c sum_f))
   (def _b_ (- avy (* _k_ avx)))
   (if (>= _b_ 0)
      (out "Function: y=" _k_ "x+" _b_)
      (out "Function: y=" _k_ "x" _b_)
   )
)
(main
   (def f_p (in "Input filePath:"))
   (def file (read f_p))
   (def items (split file (& (chr 13) (chr 10))))
   (def lst (alloc (size items)))
   (def i 0)
   (while (< i (size items))(
      (def k (split (m items i) " "))
      (array lst (i) k)
      (def i (+ i 1))
      )
   )
   (fcal lst (-(size items) 1))
)