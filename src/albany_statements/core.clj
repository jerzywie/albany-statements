(ns albany-statements.core
  (:require [dk.ative.docjure.spreadsheet :refer [select-columns load-workbook select-sheet]]
            [clojure.string :as string]
            [clojure.tools.cli :refer [parse-opts]]
            [albany-statements
             [balance :as bal]
             [util :as u]
             [config :as config]
             [sheet-status :as ss]
             [cli :as cli]
             [output :as out]])

  (:gen-class))

(def index-offset 12)
(def index-mult 2)

(defn column-index
  [[n {:keys [col]}]]
  [n (-> col (* index-mult) (+ index-offset))])

(def member-data (into {} (map column-index (:members config/config-data))))

(def common-cols {:A :code :B :description :C :case-size :E :albany-units :G :del? :H :unit-cost :K :vat-amount})

(defn gt-zero
  "Safe > 0 test. Returns falsey for any number <=0 or anything that isn't a number."
  [x] (if (instance? Number x) (pos? x) false))



(defn get-member-order
  "extracts a member's order from the sheet"
  [[name index] sheet]
  (let [m-cols (u/member-cols index [:memdes :memcost])
        order-cols (merge common-cols m-cols)
        whole-order (select-columns order-cols sheet)
        member-order (filter #(gt-zero (:memdes %))  whole-order)]
    (assoc {} name member-order)))

(defn check-sheet-status [sheet-status-closed? do-closed-check?]
  (if (not sheet-status-closed?)
    (if do-closed-check?
      (throw (IllegalStateException. "The spreadsheet must be saved and closed."))
      (println "Warning - spreadsheet is in 'Open' status - continuing anyway..."))))

(defn get-spreadsheet-data [spreadsheet-name no-closed-check?]
  (let [wb (load-workbook spreadsheet-name)
        sheet-status (ss/get-sheet-status wb)
        do-closed-check? (not no-closed-check?)
        sheet-status-closed? (= (:state sheet-status) "Closed")]
    (check-sheet-status sheet-status-closed? do-closed-check?)
    {:wb wb
     :revision (:revision sheet-status)
     :order-sheet (select-sheet "Collated Order" wb)}))

(defn generate-statements [spreadsheet-name order-date no-closed-check? suffix]
  "Do the work of statement generation."
  (let [{:keys [wb order-sheet revision]} (get-spreadsheet-data
                                           spreadsheet-name
                                           no-closed-check?)
        balance-sheet (bal/get-balance-sheet wb)
        all-orders (reduce conj (map #(get-member-order % order-sheet) member-data))]
    (println "writing statement files for " spreadsheet-name)
    (println (str "all-orders keys " (keys all-orders)))
    (println (apply str "all-orders counts " (map (fn [n] (str (name n) ":" (count (n all-orders)) " ")) (keys all-orders))))
    (doall (map #(out/emit-statement-html %
                                          all-orders
                                          (bal/get-member-balance % balance-sheet)
                                          order-date
                                          spreadsheet-name
                                          revision
                                          suffix)
                (keys member-data)))
    ))

(defn generate-orderforms [spreadsheet-name
                           order-date
                           coordinator
                           version
                           no-closed-check?
                           suffix]
  "Do the work of order-form generation."
  (let [{:keys [wb order-sheet]} (get-spreadsheet-data
                                  spreadsheet-name
                                  no-closed-check?)
        balance-sheet (bal/get-balance-sheet wb)
        all-orders (reduce conj (map #(get-member-order % order-sheet) member-data))]
    (println "writing" (string/upper-case (cli/version-tostring version)) "order forms for" spreadsheet-name)
    (println (str "all-orders keys " (keys all-orders)))
    (println (apply str "all-orders counts " (map (fn [n] (str (name n) ":" (count (n all-orders)) " ")) (keys all-orders))))
    (doall (map #(out/emit-order-html %
                                      all-orders
                                      (bal/get-member-balance % balance-sheet)
                                      order-date
                                      coordinator
                                      version
                                      suffix)
                (keys member-data)))
    ))



(defn do-generate [options-and-arguments]
  (let [{:keys [output-type
                coordinator
                version
                no-closed-check
                spreadsheet-name
                order-date
                summary
                suffix]} options-and-arguments]
    (case output-type
      :s
      (generate-statements spreadsheet-name order-date no-closed-check suffix)
      :o
      (if (nil? coordinator)
        (cli/usage summary "Coordinator name must be supplied")
        (generate-orderforms
         spreadsheet-name
         order-date
         coordinator
         version
         no-closed-check
         suffix))
      nil)))

(defn -main
  "Generate all statements from a given Albany spreadsheet."
  [& args]
  ;; work around dangerous default behaviour in Clojure
  (alter-var-root #'*read-eval* (constantly false))

  (try
    (-> args
        cli/process-args
        do-generate)
    (catch Exception e (println (str "Error: " (.getMessage e))))))
