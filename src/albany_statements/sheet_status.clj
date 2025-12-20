(ns albany-statements.sheet-status
  (:require [dk.ative.docjure.spreadsheet :refer [select-sheet select-columns]]))

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
 
