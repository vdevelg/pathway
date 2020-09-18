(defun c:asd (/ )
  (vl-load-com)

(princ "\nВыделите чертежи, предназначенные для печати")
;(setq s2 (ssget "P" '((8 . "format")) ) )

;(setq s1 (ssget  (list (assoc 8 (entget (car (entsel "Выберите объект, находящийся на необходимом слое:- "))))  )))	

(setq s1 (ssget  (list  '(8 . "format") )))	


(setq l1 (sslength s1))

(setq  blip (getvar "BLIPMODE"))
(setq  echo (getvar "CMDECHO"))
(setvar "CMDECHO"  0)
(setvar "BLIPMODE" 0)

(setq i 0) 
(while (< i= l1)

  (setq s11 (ssname s1 i))
  (setq a_nach (entget s11))
  (setq a (entget s11)) ; получение списка со всеми координатами точек прямоугольника

  	(setq t1 (assoc 10 a))    (setq t1_x (car (cdr t1))) (setq t1_y (car (cdr (cdr t1))))
	(setq a (subst '(11 11stop) t1 a))
	(setq t2 (assoc 10 a))    (setq t2_x (car (cdr t2))) (setq t2_y (car (cdr (cdr t2))))
	(setq a (subst '(11 11stop) t2 a))
	(setq t3 (assoc 10 a))    (setq t3_x (car (cdr t3))) (setq t3_y (car (cdr (cdr t3))))
	(setq a (subst '(11 11stop) t3 a))
	(setq t4 (assoc 10 a))    (setq t4_x (car (cdr t4))) (setq t4_y (car (cdr (cdr t4))))

	;(if (and (/= t1_x t2_x)(/= t1_y t2_y)) (setq tdiag_x t2_x tdiag_y t2_y )    )
	;(if (and (/= t1_x t3_x)(/= t1_y t3_y)) (setq tdiag_x t3_x tdiag_y t3_y )    )
	;(if (and (/= t1_x t4_x)(/= t1_y t4_y)) (setq tdiag_x t4_x tdiag_y t4_y )    )

	;(setq tdiag_x (max t2_x t3_x t4_x))
	;(setq tdiag_y (max t2_y t3_y t4_y))

	(setq tdiag_x_max (max t2_x t3_x t4_x))
	(setq tdiag_y_max (max t2_y t3_y t4_y))
	(setq tdiag_x_min (min t2_x t3_x t4_x))
	(setq tdiag_y_min (min t2_y t3_y t4_y))

(setq point1 (list tdiag_x_min tdiag_y_min))
(setq point2 (list tdiag_x_max tdiag_y_max))

(setq Dy (- tdiag_y_max tdiag_y_min))
(setq Dx (- tdiag_x_max tdiag_x_min))


;(setq point1 (list t1_x t1_y))
;(setq point2 (list tdiag_x tdiag_y))

;(setq x1 (min tdiag_x t1_x)) (setq x2 (max tdiag_x t1_x))
;(setq y1 (min tdiag_y t1_y)) (setq y2 (max tdiag_y t1_y)) 

;(setq Dy (- y2 y1))
;(setq Dx (- x2 x1))


(setq Dyx (/ Dy Dx))

(if (> Dy Dx ) (setq ugol "_P")(setq ugol "_l"))

(command "_-plot" "_y"  ""  ""  ""  ""  ugol  ""   "" point1 point2 ""  ""  ""   ""  ""  ""   ""  ""   "")
(setq i (1+ i))

)

(setvar "CMDECHO"  echo)
(setvar "BLIPMODE" blip)

)
