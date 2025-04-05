(fn l99()
   (def x 1)
   (def crlf (chr 13))
   (def fom "")
   (while (<= x 9)(
      (def y 1)
      (while(<= y x)(
         (def fom (& fom y "*" x "=" (* y x) " "))
         (def y (+ y 1))
         )
      )
      (def fom (& fom crlf))
      (def x (+ x 1))
      )
   )
   (out fom)
)

(main
   (l99())
)