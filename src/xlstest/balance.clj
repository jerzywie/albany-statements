(ns xlstest.balance
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:gen-class)
  (:require [clojure.string :as s]
            [xlstest.util :as util]
            [hiccup.core :as h]
            [hiccup.page :as p]
            [garden.core]
            [garden.stylesheet]))



(def member-data {:sally 4 :isabel 6 :jerzy 8 :alice 10 :carol 12 :clair 14 :annabelle 16 :deborah 18 :frank 20 :matthew 22})

(def common-cols {:A :name :C :subname :Z :key})

(def bal-keys [:bf :ordertotal :joinfee :levy :owed :moneyin :balance :cooptomember :membertocoop :cf])

(def optional-bal-keys [:membertocoop :joinfee :levy :cooptomember])

(defn get-balance-sheet
  [wb]
  (select-sheet "Financials" wb))

(defn keywordize-bal [b]
  {(keyword (:key b)) (dissoc b :key)})

(defn get-member-balance
  "extract a member's balance information into a map of maps"
  [name sheet]
  (let [cols (merge common-cols (util/member-cols (name member-data) [:val]))
        raw-bal (select-columns cols sheet)
        clean-bal (map keywordize-bal ( filter #(:key %) raw-bal))]
    (apply merge clean-bal)))

(defn bal-item
  "produce name value pair for balance item,
   unless it's an optional item with value zero"
  [k bal]
  (let [item (k bal)
        val (:val item)
        finalval (if (nil? val) 0.0 val)
        name (s/trim (str (:name item) " " (:subname item))) ]
    (when-not (and (some #{k} optional-bal-keys)
                    (= finalval 0.0))
      [name finalval])))
