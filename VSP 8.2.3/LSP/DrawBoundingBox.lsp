(defun c:DrawBoundingBox (/ ss minpt maxpt obj x y z)
  (vl-load-com) ; Load the Visual LISP COM interface
  (prompt "\nSelect objects to create bounding box: ")
  
  ; Get a selection set from the user
  (setq ss (ssget))
  
  ; Only proceed if the selection is not nil
  (if ss
    (progn
      ; Initialize the min and max points
      (setq minpt (list 1e99 1e99 0))
      (setq maxpt (list -1e99 -1e99 0))
      
      ; Loop through all the objects in the selection set
      (foreach ename (mapcar 'cadr (ssnamex ss))
        (setq obj (vlax-ename->vla-object ename)) ; Convert entity name to a VLA-object
        
        ; Get the object's geometric extents (bounding box)
        (vl-catch-all-apply 'vla-getboundingbox (list obj 'x 'y))
        (setq x (vlax-safearray->list x))
        (setq y (vlax-safearray->list y))
        
        ; Update min and max points
        (setq minpt (mapcar 'min minpt x))
        (setq maxpt (mapcar 'max maxpt y))
      )
      
      ; Draw a rectangle using the min and max points to define the corners
      (command "_RECTANGLE" minpt maxpt)
    )
    ; Inform the user if nothing was selected
    (prompt "\nNo objects selected.")
  )
  ; Clean up and return to the command prompt
  (princ)
)

; Inform the user how to run the command after loading the LISP file
(princ "\nType DrawBoundingBox to create a bounding box around selected objects.")