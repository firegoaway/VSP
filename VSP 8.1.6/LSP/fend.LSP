(defun c:fend (/ l1 l2 s1 s2
               p10 p11 p20 p21 e1 e2 p1 p2 em1 em2 la
               sp ep sa ea mp ra pt1 pt2 pp1 pp2 zl pe1 pe2
               )

 (princ "\nSelect 2 Parallel Lines")
 (while (not l1)
        (while (not s1)
               (setq s1 (entsel "\nSelect 1st Line:   ")))
        (setq p1 (nth 1 s1)
              ep (car s1))
        (if (= "LINE" (cdr (assoc 0 (entget ep))))
            (setq l1 ep)))
 (setq pe1 (osnap p1 "end"))
 (redraw l1 3)

 (while (not l2)
        (while (not s2)
               (setq s2 (entsel "\nSelect 2nd Line:   ")))
        (setq p2 (nth 1 s2)
              ep (car s2))
        (if (= "LINE" (cdr (assoc 0 (entget ep))))
            (setq l2 ep)))
 (setq pe2 (osnap p2 "end"))
 (redraw l2 3)

 (setq e1 (entget l1)
       e2 (entget l2)
       p10 (cdr (assoc 10 e1))
       p11 (cdr (assoc 11 e1))
       p20 (cdr (assoc 10 e2))
       p21 (cdr (assoc 11 e2)))

;;;COMPARE LINE ANGLES
 (if (not (or (equal (angle p10 p11) (angle p20 p21) 1e-
              (equal (angle p10 p11) (angle p21 p20) 1e-))
      (progn
        (alert "Lines Are Not Parallel")
        (exit)))

;;;COMPARE LINE ELEVATIONS
 (setq zl (mapcar 'caddr (list p10 p11 p20 p21)))
 (if (apply '/= zl)
     (progn
        (alert "Lines Are Not Equal Elevation")
        (exit)))

;;;COMPARE LINE UCS
 (if (not (equal (cdr (assoc 210 e1)) (cdr (assoc 210 e2)) 1e-)
     (alert "Lines Are Not Same UCS - Be Careful"))

;;;CLOSEST ENDS TO 1ST PICK POINT
 (if (< (distance p10 p1) (distance p11 p1))
     (setq pt1 p10 em1 10 la (angle p11 p10))
     (setq pt1 p11 em1 11 la (angle p10 p11)))
 (if (< (distance p20 pt1) (distance p21 pt1))
     (setq pt2 p20 em2 10)
     (setq pt2 p21 em2 11))
 (grdraw pt1 pt2 3 0)

;;;PERPINDICULAR POINTS
 (setq pp1 (inters p10 p11 pt2 (polar pt2 (+ (angle p10 p11) (* pi 0.5)) 1) nil))
 (setq pp2 (inters p20 p21 pt1 (polar pt1 (+ (angle p10 p11) (* pi 0.5)) 1) nil))
 (grdraw pt1 pp2 1 0)
 (grdraw pt2 pp1 2 0)

 (setq mp (mapcar '(lambda (a b) (* (+ a b) 0.5)) pt1 pt2)
       ra (* (distance pt1 pp2) 0.5)
       sa (angle pt1 pp2)     ;;;!!!Here is the problem child
       ea (+ sa pi)
       sp (polar mp sa ra)
       ep (polar mp ea ra))

 (entmake (list (cons 0 "ARC")
                (assoc 8 e1)
                (cons 10 mp)
                (cons  6 (if (assoc  6 e1) (cdr (assoc  6 e1)) "BYLAYER"))
                (cons 39 (if (assoc 39 e1) (cdr (assoc 39 e1)) 0))
                (cons 62 (if (assoc 62 e1) (cdr (assoc 62 e1)) 256))
                (cons 40 ra)
                (cons 50 sa)
                (cons 51 ea)))

 (grdraw sp mp 4 0)
 (grdraw ep mp 5 0)

;;;TRIM THE LINES
 (entmod (subst (cons em1 ep) (assoc em1 e1) e1))
 (entmod (subst (cons em2 sp) (assoc em2 e2) e2))

;;FIX THEN BUG
 (initget "Yes No")
 (if (= "Yes" (getkword "\nRotate The ARC 180 Degrees? <N>:   "))
     (command "_.ROTATE" (entlast) "" mp 180.0))

 (redraw)

 (prin1))