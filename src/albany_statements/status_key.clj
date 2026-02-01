(ns albany-statements.status-key
  (:require [dk.ative.docjure.spreadsheet :refer [select-sheet select-columns]]
            [clojure.string :as s]
            [albany-statements.util :as util]))

(def cols {:A :code :C :description})

(defn get-status-codes-and-text
  [wb]
  (->> wb (select-sheet "ND Lookup") (select-columns cols)))
