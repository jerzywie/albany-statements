(ns albany-statements.balance
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:gen-class)
  (:require [clojure.string :as s]
            [albany-statements
             [util :as util]
             [config :as config]]
            [hiccup.core :as h]
            [hiccup.page :as p]
            [garden.core]
            [garden.stylesheet]))


(def index-offset 2)
(def index-mult 2)
(def key-offset-from-last-member 3)

(defn column-index
  [[n {:keys [col]}]]
  [n (-> col (* index-mult) (+ index-offset))])

(def member-data (into {} (map column-index (:members config/config-data))))

(def key-column (->> (vals member-data) (apply max) (+ key-offset-from-last-member) util/to-col))

(def common-cols {:A :name :C :subname key-column :key})

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
