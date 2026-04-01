;;; ==========================================================================
;;; 3_BeltLine.lsp (Command: BELT)
;;; --------------------------------------------------------------------------
;;; 機能: 2つの円を選択し、それらを囲むベルト状のポリラインを作成する
;;; --------------------------------------------------------------------------
;;; Target: AutoCAD 2024 and earlier (No LT) / AutoCAD 2024以下 (LT不可)
;;; Encoding: UTF-8 with BOM (Must be preserved / 保存時は必ずBOM付きUTF-8を指定すること)
;;; ==========================================================================

;; Note: Top-level vl-load-com removed to enforce guard check inside command
;; 注: ガードチェックを強制するため、トップレベルの vl-load-com は削除済み

(defun c:BELT ( /
               ;; Local Variables (All pointers must be local / 全てローカル変数化)
               doc
               sysvars
               restoreFailList
               undoStarted
               interactiveMode
               ;; Helper Functions
               internal:save-sysvars
               internal:restore-sysvars
               internal:cleanup
               internal:get-circle-geom
               internal:mxp
               internal:acos
               *error*
               ;; Command Variables
               sel1 sel2 data1 data2 c1 r1 c2 r2
               dist ang ang-diff p1-a p1-b p2-a p2-b
             )

  ;; -------------------------------------------------------------------------
  ;; Initialize Variables (Explicit initialization / 変数の明示的初期化)
  ;; -------------------------------------------------------------------------
  (setq restoreFailList nil
        undoStarted     nil
        interactiveMode nil)

  ;; -------------------------------------------------------------------------
  ;; 0. Check ActiveX (Absolute Requirement / ActiveX必須チェック)
  ;; -------------------------------------------------------------------------
  (if (vl-catch-all-error-p (vl-catch-all-apply 'vl-load-com nil))
    (progn
      (if (not (member "ActiveX" restoreFailList))
        (setq restoreFailList (cons "ActiveX" restoreFailList)))
      (princ "\n[CRITICAL WARNING] vl-load-com failed / vl-load-comに失敗しました")
      (exit)
    )
  )

  ;; -------------------------------------------------------------------------
  ;; 1. Get Document Object (Doc取得)
  ;; -------------------------------------------------------------------------
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (if (not doc)
    (progn
      (if (not (member "Doc" restoreFailList))
        (setq restoreFailList (cons "Doc" restoreFailList)))
      (princ "\n[CRITICAL WARNING] Failed to get Document / ドキュメントの取得に失敗しました")
      (exit)
    )
  )

  ;; -------------------------------------------------------------------------
  ;; 2. Start Undo Mark (Undo開始)
  ;; -------------------------------------------------------------------------
  (if (vl-catch-all-error-p (vl-catch-all-apply 'vla-StartUndoMark (list doc)))
    (progn
      (if (not (member "Undo" restoreFailList))
        (setq restoreFailList (cons "Undo" restoreFailList)))
      (princ "\n[CRITICAL WARNING] Failed to start Undo / Undoの開始に失敗しました")
      (exit)
    )
    (setq undoStarted T)
  )

  ;; -------------------------------------------------------------------------
  ;; System Variables Definition (List of lists / リストのリスト形式)
  ;; Format: (("VARNAME" Value) ...) - Dot pairs are PROHIBITED / ドット対禁止
  ;; -------------------------------------------------------------------------
  (setq sysvars '(("CMDECHO" 0) ("OSMODE" 0)))

  ;; -------------------------------------------------------------------------
  ;; Helper: Save System Variables (PNR Trigger / ここからPNR確定)
  ;; -------------------------------------------------------------------------
  (defun internal:save-sysvars ( / new-list original-val pair)
    (setq new-list nil)
    (foreach pair sysvars
      (setq original-val (getvar (car pair)))
      (if original-val
        (progn
          (setq new-list (cons (list (car pair) original-val) new-list))
          (if (vl-catch-all-error-p (vl-catch-all-apply 'setvar pair))
            (princ (strcat "\n[WARNING] Failed to set sysvar: " (car pair)))
          )
        )
      )
    )
    (setq sysvars new-list)
    (princ)
  )

  ;; -------------------------------------------------------------------------
  ;; Helper: Restore System Variables
  ;; -------------------------------------------------------------------------
  (defun internal:restore-sysvars ( / pair)
    (foreach pair sysvars
      (if (vl-catch-all-error-p (vl-catch-all-apply 'setvar pair))
        (if (not (member (car pair) restoreFailList))
          (setq restoreFailList (cons (car pair) restoreFailList)))
      )
    )
    (princ)
  )

  ;; -------------------------------------------------------------------------
  ;; Helper: Cleanup & Final Report
  ;; -------------------------------------------------------------------------
  (defun internal:cleanup ()
    (internal:restore-sysvars)
    (if undoStarted (vla-EndUndoMark doc))
    (if restoreFailList
      (progn
        (if (or (member "ActiveX"  restoreFailList)
                (member "Doc"      restoreFailList)
                (member "Undo"     restoreFailList)
                (member "UCS"      restoreFailList)
                (member "UCS-SAVE" restoreFailList))
          (princ "\n[CRITICAL WARNING] ")
          (princ "\n[WARNING] ")
        )
        (princ
          (vl-string-trim ", "
            (apply 'strcat
              (mapcar '(lambda (x) (strcat x ", ")) restoreFailList)
            )
          )
        )
      )
    )
    (princ)
  )

  ;; -------------------------------------------------------------------------
  ;; Helper: Matrix × Point (変換行列 × 点)
  ;; -------------------------------------------------------------------------
  (defun internal:mxp (m p / x y z r0 r1 r2 tr)
    (setq x  (car p)    y  (cadr p)  z  (caddr p))
    (if (null z) (setq z 0.0))
    (setq r0 (car m)    r1 (cadr m)  r2 (caddr m)  tr (cadddr m))
    (list
      (+ (* x (car r0))   (* y (car r1))   (* z (car r2))   (car tr))
      (+ (* x (cadr r0))  (* y (cadr r1))  (* z (cadr r2))  (cadr tr))
      (+ (* x (caddr r0)) (* y (caddr r1)) (* z (caddr r2)) (caddr tr))
    )
  )

  ;; -------------------------------------------------------------------------
  ;; Helper: Arc Cosine (逆余弦)
  ;; -------------------------------------------------------------------------
  (defun internal:acos (x)
    (cond
      ((>= x  1.0)  0.0)
      ((<= x -1.0)  pi)
      (t (atan (sqrt (- 1.0 (* x x))) x))
    )
  )

  ;; -------------------------------------------------------------------------
  ;; Helper: Get Circle Geometry (円ジオメトリ取得、nentsel変換行列対応)
  ;; -------------------------------------------------------------------------
  (defun internal:get-circle-geom (sel-list / ent data cen rad mat scale)
    (setq ent  (car sel-list)
          data (entget ent)
          cen  (cdr (assoc 10 data))
          rad  (cdr (assoc 40 data)))
    (if (= (length sel-list) 4)
      (progn
        (setq mat   (caddr sel-list)
              scale (distance (internal:mxp mat '(0 0 0)) (internal:mxp mat '(1 0 0))))
        (setq rad (* rad scale))
        (setq cen (internal:mxp mat cen))
      )
    )
    (list cen rad)
  )

  ;; -------------------------------------------------------------------------
  ;; Error Handler
  ;; -------------------------------------------------------------------------
  (defun *error* (msg)
    (if (and msg (not (wcmatch (strcase msg) "*CANCEL*,*QUIT*,*BREAK*")))
      (princ (strcat "\n[Error] " msg))
    )
    (internal:cleanup)
    (princ)
  )

  ;; -------------------------------------------------------------------------
  ;; 3. Point of No Return (Sysvar Save) / ここから復元義務発生
  ;; -------------------------------------------------------------------------
  (internal:save-sysvars)

  ;; -------------------------------------------------------------------------
  ;; Main Logic
  ;; -------------------------------------------------------------------------
  (setq interactiveMode T)

  ;; 1つ目の円を選択
  (while (and (null sel1) (not (equal (getvar "LASTPROMPT") "1つ目の円を選択: ")))
    (setq sel1 (nentsel "\n1つ目の円を選択: "))
    (if (and sel1 (/= (cdr (assoc 0 (entget (car sel1)))) "CIRCLE"))
      (progn (princ "\n円ではありません。") (setq sel1 nil))
    )
  )

  ;; 2つ目の円を選択
  (if sel1
    (while (and (null sel2) (not (equal (getvar "LASTPROMPT") "2つ目の円を選択: ")))
      (setq sel2 (nentsel "\n2つ目の円を選択: "))
      (cond
        ((null sel2) nil)
        ((/= (cdr (assoc 0 (entget (car sel2)))) "CIRCLE")
         (princ "\n円ではありません。") (setq sel2 nil))
        ((and (equal (car sel1) (car sel2))
              (or (< (length sel1) 4)
                  (equal (cadddr sel1) (cadddr sel2))))
         (princ "\n同じ円です。") (setq sel2 nil))
      )
    )
  )

  (if (and sel1 sel2)
    (progn
      (setq data1 (internal:get-circle-geom sel1)
            data2 (internal:get-circle-geom sel2)
            c1    (car data1)  r1 (cadr data1)
            c2    (car data2)  r2 (cadr data2)
            dist  (distance c1 c2))

      (cond
        ((equal dist 0.0 1e-4)
         (princ "\n[エラー] 中心点が同じため作成できません。"))

        ((<= dist (abs (- r1 r2)))
         (princ "\n[エラー] 円が近すぎるか、一方が他方を内包しているため作成できません。"))

        (t
         (setq ang      (angle c1 c2)
               ang-diff (internal:acos (/ (- r1 r2) dist))
               p1-a     (polar c1 (+ ang ang-diff) r1)
               p2-a     (polar c2 (+ ang ang-diff) r2)
               p1-b     (polar c1 (- ang ang-diff) r1)
               p2-b     (polar c2 (- ang ang-diff) r2))

         (vl-cmdf "_.PLINE" p1-a p2-a "_A" p2-b "_L" p1-b "_A" "_CL")
        )
      )
    )
    (prompt "\n選択がキャンセルされました。")
  )

  ;; -------------------------------------------------------------------------
  ;; 4. Normal Exit / 正常終了 (Silent Success / 成功時無言)
  ;; -------------------------------------------------------------------------
  (internal:cleanup)
  (princ)
)

(princ "\n[OK] BELT コマンドがロードされました。")
(princ)
