(ns albany-statements.core
  (:require [dk.ative.docjure.spreadsheet :refer [select-columns load-workbook select-sheet]]
            [clojure.string :as string]
            [hiccup.core :as h]
            [hiccup.page :as p]
            [garden.core]
            [clojure.tools.cli :refer [parse-opts]]
            [garden.stylesheet]
            [albany-statements
             [balance :as bal]
             [util :as u]
             [config :as config]
             [sheet-status :as ss]
             [cli :as cli]])

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

(defn member-display-name [name-keyword] (clojure.string/capitalize (name name-keyword)))

(defn get-member-order
  "extracts a member's order from the sheet"
  [[name index] sheet]
  (let [m-cols (u/member-cols index [:memdes :memcost])
        order-cols (merge common-cols m-cols)
        whole-order (select-columns order-cols sheet)
        member-order (filter #(gt-zero (:memdes %))  whole-order)]
    (assoc {} name member-order)))

;; start of html and css-realted functions
(def dark-blue "#A4D1FE")
(def light-blue "#EAF4FF")
(def light-grey "#ECECEA")

(defn- fix-css-fn [s]
  "An ugly hack to get round Garden's lack of complete CSS3 function support."
  (clojure.string/replace (clojure.string/replace s "<" "(") ">" ")"))

(defn gen-statement-css []
  (fix-css-fn
   (garden.core/css
    [:body
     {:font-family "Verdana, Arial, Helvetica, sans-serif"}
     {:font-size "12px"}
     {:font-weight "normal"}
     ]
    [:.order :.balance {:border-collapse "collapse"}
     [:td :th {:border "#999 1px solid"} ]
     [:td {:padding "4px"} ]
     [:th {:background  dark-blue} {:padding "8px"}]
     [:tr
      [:&:nth-child<even> {:background light-grey}]
      [:&:nth-child<odd> {:background light-blue}]]
     ]
    [:table.order {:width "100%"}]
    [:div.balance {:float "right"} {:margin-top "8px" :margin-bottom "8px"}]
    [:.rightjust {:text-align "right"}]
    [:.bold {:font-weight "bold"}]
    )))

(defn emit-balance-html
  [key balance]
  (let [[n v :as item] (bal/bal-item key balance)]
    (when-not (nil? item)
      [:tr
       [:td n]
       [:td.rightjust (u/tocurrency v)]
       ])))

(defn statement-head [member-name]
  [:head
   [:title (str "Statement for " (member-display-name member-name))]
   [:meta {:charset "UTF-8"}]
   [:style {:type "text/css"} (gen-statement-css)]])

(defn statement-body [member-name
                      member-order
                      balance
                      order-date
                      order-total
                      spreadsheet-name
                      revision]
  [:body
   [:h1 "Albany Coop"]
   [:h2 (str "Statement for " (member-display-name member-name) " - " order-date)]
   [:div.order
    [:table.order
     [:thead
      [:tr
       [:th {:rowspan 2} "Code"]
       [:th {:rowspan 2} "Description"]
       [:th {:rowspan 2} "Case"]
       [:th {:colspan 3} "Cost per Case"]
       [:th {:rowspan 2} "Albany units/case"]
       [:th {:colspan 2} "Amount purchased"]
       [:th {:rowspan 2}"Status"]
       [:th {:rowspan 2}"Price charged"]]
      [:tr
       [:th "Net"]
       [:th "VAT"]
       [:th "Gross"]
       [:th "Albany units"]
       [:th "Cases"]
       ]]
     (into [:tbody]
           (for [line member-order]
             [:tr
              [:td (:code line)]
              [:td (u/format-description (:description line))]
              [:td (:case-size line)]
              [:td.rightjust (u/tocurrency (:unit-cost line))]
              [:td.rightjust (u/tocurrency (:vat-amount line))]
              [:td.rightjust (u/tocurrency (+ (:unit-cost line) (:vat-amount line)))]
              [:td.rightjust (:albany-units line)]
              [:td.rightjust (:memdes line)]
              [:td.rightjust (u/essential-cases (:memdes line) (:albany-units line))]
              [:td (:del? line)]
              [:td.rightjust (u/tocurrency (:memcost line))]
              ]
             ))
     [:tbody
      [:tr.bold
       [:td {:colspan 10} (str "Total for " order-date)]
       [:td.rightjust (u/tocurrency order-total)]
       ]
      ]
     ]
    ]
   [:div.balance
    [:table.balance
      [:thead
       [ :tr
        [:th {:colspan 2} "Balance information"]
        ]]
      [:tbody
       (map #(emit-balance-html % balance) bal/bal-keys)
       ]
     ]
    ]
   [:div
    [:table.order
     [:tbody
      [:tr
       [:td (format "Taken from %s. Save revision %d" spreadsheet-name (int revision))]]]]]])


(defn gen-order-css []
  (fix-css-fn
   (garden.core/css
    [:body
     {:font-family "Verdana, Arial, Helvetica, sans-serif"}
     {:font-size "12px"}
     {:font-weight "normal"}
     ]
    [:.order {:width "100%"} {:border-collapse "collapse"}
     [:td :th {:border "#999 1px solid"} ]
     [:td {:padding "4px"} ]
     [:th {:background  dark-blue} {:padding "8px"}]
     ]
    [:.rightjust {:text-align "right"}]
    [:td.space {:padding-top "30px"}]
    [:.bigtext {:font-size "1.5em"}]
    (garden.stylesheet/at-media {:print true}
                                [:thead {:display "table-header-group"}])
    )))

(defn order-head [member-name]
  [:head
   [:title (str "Order for " (member-display-name member-name))]
   [:meta {:charset "UTF-8"}]
   [:style {:type "text/css"} (gen-order-css)]])

(defn order-body [member-name member-order order-date order-total coordinator version]
  [:body
   [:h1 "Albany Food Coop     Order Form"]
   (case version
     :d [:h2 "Pre-meeting DRAFT"]
     :f [:h2 "FINAL for order sorting"])
   [:table.order
    [:thead
     [:tr
      [:th [:span "Name: " [:b.bigtext (member-display-name member-name)]]]
      [:th [:span "Order Date: " [:b.bigtext order-date]]]
      [:th (str "Coordinator: " coordinator)]
      ]]]
   [:p " "]
   [:table.order
    [:thead
     [:tr
      [:th {:rowspan 2} "Code"]
      [:th {:rowspan 2} "Item"]
      [:th {:rowspan 2} "Case"]
      [:th {:rowspan 2} "Unit Price"]
      [:th {:rowspan 2} "Albany units/case"]
      [:th {:colspan 2} "Pref Q'ty"]
      [:th {:colspan 2} "Actual Q'ty"]
      [:th {:rowspan 2} "Est price"]
      [:th {:rowspan 2} "Final Q'ty"]
      ]
     [:tr
      [:th "Albany units"]
      [:th "Cases"]
      [:th "Albany units"]
      [:th "Cases"]
      ]
     ]
    (into [:tbody]
          (for [line member-order]
            [:tr
             [:td [:b (:code line)]]
             [:td (u/format-description (:description line))]
             [:td (:case-size line)]
             [:td.rightjust (u/tocurrency (:unit-cost line))]
             [:td.rightjust (:albany-units line)]
             [:td.rightjust (:memdes line)]
             [:td.rightjust (u/essential-cases (:memdes line) (:albany-units line))]
             [:td " "]
             [:td " "]
             [:td.rightjust (u/tocurrency (:memcost line))]
             [:td.rightjust (if-not (nil? (:del? line)) (str "(" (:del? line) ")") " ")]]
            ))
    [:tbody
     [:tr
      [:td.rightjust {:colspan 9} [:b "Estimated Sub-total for above items "]]
      [:td.rightjust [:b (u/tocurrency order-total)]]
      [:td " "]
      ]]

    (into [:tbody]
          (repeat 10
                  [:tr (repeat 11 [:td.space " "])]))
    ]
   ])

(defn emit-statement-html [member-name
                           all-orders
                           mem-balance
                           order-date
                           spreadsheet-name
                           revision]
  (let [member-order (member-name all-orders)
        order-total (reduce #(+ %1 (:memcost %2)) 0 member-order)
        file-name (str (member-display-name member-name) "-" order-date ".html")]
    (println (str "File-name " file-name " Order-total " (format "%.2f" (double order-total)) ))
    (spit  file-name
           (p/html5 (statement-head member-name)
                    (statement-body member-name
                                    member-order
                                    mem-balance
                                    order-date
                                    order-total
                                    spreadsheet-name
                                    revision)))))

(defn emit-order-html [member-name all-orders order-date coordinator version]
  (let [member-order (member-name all-orders)
        order-total (reduce #(+ %1 (:memcost %2)) 0 member-order)
        file-name (str (member-display-name member-name) "-" order-date "-OrderForm-" (string/upper-case (cli/version-tostring version)) ".html")]
    (println (str "File-name " file-name " Order-total " (format "%.2f" (double order-total)) ))
    (spit  file-name
           (p/html5 (order-head member-name)
                    (order-body member-name
                                member-order
                                order-date
                                order-total
                                coordinator
                                version)))))

;; end of html and css-related stuff

(defn get-spreadsheet-data [spreadsheet-name]
  (let [wb (load-workbook spreadsheet-name)
        sheet-status (ss/get-sheet-status wb)]
    (when (not= (:state sheet-status) "Closed")
      (throw (IllegalStateException. "The spreadsheet must be saved and closed.")))
    {:wb wb
     :revision (:revision sheet-status)
     :order-sheet (select-sheet "Collated Order" wb)}))

(defn generate-statements [spreadsheet-name order-date]
  "Do the work of statement generation."
  (let [{:keys [wb order-sheet revision]} (get-spreadsheet-data spreadsheet-name)
        balance-sheet (bal/get-balance-sheet wb)
        all-orders (reduce conj (map #(get-member-order % order-sheet) member-data))]
    (println (str "all-orders keys " (keys all-orders)))
    (println (apply str "all-orders counts " (map (fn [n] (str (name n) ":" (count (n all-orders)) " ")) (keys all-orders))))
    (doall (map #(emit-statement-html %
                                      all-orders
                                      (bal/get-member-balance % balance-sheet)
                                      order-date
                                      spreadsheet-name
                                      revision)
                (keys member-data)))
    ))

(defn generate-orderforms [spreadsheet-name order-date coordinator version]
  "Do the work of order-form generation."
  (let [{:keys [wb order-sheet]} (get-spreadsheet-data spreadsheet-name)
        all-orders (reduce conj (map #(get-member-order % order-sheet) member-data))]
    (println (str "all-orders keys " (keys all-orders)))
    (println (apply str "all-orders counts " (map (fn [n] (str (name n) ":" (count (n all-orders)) " ")) (keys all-orders))))
    (doall (map #(emit-order-html % all-orders order-date coordinator version) (keys member-data)))
    ))



(defn do-generate [options-and-arguments]
  (let [{:keys [output-type
                coordinator
                version
                spreadsheet-name
                order-date
                summary]} options-and-arguments]
    (case output-type
      :s
      (do (println "writing statement files for " spreadsheet-name)
          (generate-statements spreadsheet-name order-date))
      :o
      (if (nil? coordinator)
        (cli/usage summary "Coordinator name must be supplied")
        (do (println "writing" (string/upper-case (cli/version-tostring version)) "order forms for" spreadsheet-name)
            (generate-orderforms spreadsheet-name order-date coordinator version)))
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
