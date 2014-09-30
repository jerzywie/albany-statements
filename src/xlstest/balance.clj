(ns xlstest.balance
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:gen-class)
  (:require [xlstest.util :as util]
            [hiccup.core :as h]
            [hiccup.page :as p]
            [garden.core]
            [garden.stylesheet]))



(def member-data {:sally 13 :isabel 15 :jerzy 17 :alice 19 :carol 21 :clair 23 :annabelle 25 :deborah 27 :ann 29 :matthew 31})

(defn do-balances [spreadsheet-name]
  (let [wb (load-workbook spreadsheet-name)
        bal-sheet (select-sheet "Financials" wb)
        all-orders (reduce conj (map #(get-member-order % order-sheet) member-data))))
