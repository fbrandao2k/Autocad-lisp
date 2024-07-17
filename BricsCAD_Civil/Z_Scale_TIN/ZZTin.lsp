; ZZTIN
;
; CAD Concepts
; 2021-09-24
; https://www.cadconcepts.co.nz
;
; Scales a copy of selected BricsCAD TIN Surfaces by the given scale factor.
; The scaled TIN surface put on its on layer named after the parent object suffixed with the scale factor.

; Requires BricsCAD Pro V21 or greater.
; It could be adjusted to work with BricsCAD V20 Platinum, which included Civil tools

(defun C:ZZTin ( / sset oDoc ans TIN ZTIN LayName oLayers TIN_pts TIN_pts2)
	(cond ((and (eq (getvar 'PROGRAM) "BRICSCAD") (>= (atoi (car (vle-string-split "." (getvar '_VERNUM)))) 21) (= 4 (boole 1 (getvar 'LICFLAGS) 4))) ; Check to see if it's BricsCAD Pro V20 or greater
				(vl-load-tin) ; Load TIN surface extension
					(princ "\nSelect TIN surfaces to work with ")
					(setq sset (ssget '((0 . "BsysCvDbTinSurface"))))
					(cond (sset
									(setq oDoc (vla-get-activedocument (vlax-get-acad-object))); retrieve current doc
				   				(vla-EndUndoMark oDoc)
									(vla-StartUndoMark oDoc)
									(if (not *zscale*) (setq *zscale* 1.0)) ; set default scale
									(initget 2); disallow 0 input
									(if (setq ans (getdist (strcat "\nEnter a Z scale factor: <" (rtos *zscale* 2 3) ">: ")))
 											(setq *zscale* ans)
 									)
									(foreach TIN (vle-selectionset->list  sset)
				 						(setq ZTIN (vla-copy (vlax-ename->vla-object TIN))) ; create a copy
				 						(setq LayName (strcat (vla-get-layer ZTIN) "-ZZTIN-" (rtos *zscale* 2 3))) ; create target layer name
				 						(setq oLayers (vla-get-layers oDoc)); retrieve layers
										;Check that our given layer exists and create it if it doesn't
										(if (vl-catch-all-error-p 
													(vl-catch-all-apply 'vla-item (list oLayers layname))
												)
												(vla-add oLayers LayName) ; create the layer
										)
				 						(vla-put-layer ZTIN layname) ; Put the copied TIN on the target layer
										(setq TIN_pts (tin:getpoints ZTIN T)) ; Get existing TIN points
										;---
										; Method to change elevation using tin:movepoints
										;(setq Tin_pts2 (mapcar '(lambda (pt) (mapcar '* (list 1.0 1.0 *zscale*) pt)) Tin_pts)) ; Calculate target Tin points
										;(tin:movepoints ZTIN TIN_pts TIN_pts2)
										;---
										; Method to change elevation using changePointElevations
										(setq Tin_pts2 (mapcar '(lambda (pt) (* *zscale* (caddr pt))) Tin_pts)) ; Calculate just the target Z values
										(tin:changePointElevations ZTIN TIN_pts TIN_pts2) ; Modify the TIN
										;---
										(tin:updatetin ZTIN) ; Force an update
									)
									(vla-EndUndoMark oDoc)
								)
								(T (princ"\nNo TIN surfaces selected"))
					)
				)
				(T (Princ "\nCivil Tools require BricsCAD Pro 21 or greater")) ; Display a warning on the command line if its not
	)
	(prin1)
)

(princ "\nZZTin Loaded. Command(s): ZZTIN.")
(prin1)