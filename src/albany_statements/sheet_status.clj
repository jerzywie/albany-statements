(ns albany-statements.sheet-status
  (:use [dk.ative.docjure.spreadsheet] :reload-all))

(def status-sheet "Save Revision")
(def cols {:A :item :B :value})

(defn- open-status-sheet
  [wb]
  (select-sheet status-sheet wb))

(defn get-sheet-status
  [wb]
  (let [sheet (open-status-sheet wb)
        raw-status (drop 2 (select-columns cols sheet))]
    (->> raw-status
        (map (fn [m] {(keyword (:item m)) (:value m)}))
        (apply merge))))
 
