(ns xlstest.core
  (:use [dk.ative.docjure.spreadsheet] :reload-all)
  (:gen-class)
  (:require [hiccup.core :as h])
  (:require [hiccup.page :as p])
  (:require [garden.core])
  (:require [garden.stylesheet]))


;; Random notes and snippets for pulling individual members orders from the spreadsheet


;;This was obtained from Stackoverflow
(defn to-col [num]
  (loop [n num s ()]
    (if (> n 25)
      (let [r (mod n 26)]
        (recur (dec (/ (- n r) 26)) (cons (char (+ 65 r)) s)))
      (keyword (apply str (cons (char (+ 65 n)) s))))))

;; this produces the map of columns for a given member
(defn member-cols
  "Produce the map of columns for a given member"
  [index]
 (zipmap  (vec  (map to-col (range index (+ 2 index)))) [:memdes :memcost] ))

;; it can then be merge'd with the common-cols

(def member-data {:sally 13 :isabel 15 :jerzy 17 :alice 19 :carol 21 :clair 23 :annabelle 25 :deborah 27 :ann 29 :matthew 31})

(def common-cols {:A :code :B :description :C :case-size :E :albany-units :G :del? :H :unit-cost :J :vat-amount})

(defn gt-zero
  "Safe > 0 test. Returns falsey for any number <=0 or anything that isn't a number."
  [x] (if (instance? Number x) (pos? x) false))

(defn member-display-name [name-keyword] (clojure.string/capitalize (name name-keyword)))

(defn get-member-order
  "extracts a member's order from the sheet"
  [[name index] sheet]
  (let [m-cols (member-cols index)
        order-cols (merge common-cols m-cols)
        whole-order (select-columns order-cols sheet)
        member-order (filter #(gt-zero (:memdes %))  whole-order)]
    (assoc {} name member-order)))

(defn tocurrency
  "Format a number for currency display (2 digits. Parentheses when negative)"
  [v]
  (format "%.2f" (double v)))

;; start of html and css-realted functions
(def dark-blue "#A4D1FE")
(def light-blue "#EAF4FF")

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
    [:.order {:width "100%"} {:border-collapse "collapse"}
     [:td :th {:border "#999 1px solid"} ]
     [:td {:padding "4px"} ]
     [:th {:background  dark-blue} {:padding "8px"}]
     [:tr
      [:&:nth-child<even> {:background dark-blue}]
      [:&:nth-child<odd> {:background light-blue}]]
     ]
    [:.rightjust {:text-align "right"}]
    )))



(defn statement-head [member-name]
  [:head
   [:title (str "Statement for " (member-display-name member-name))]
   [:style {:type "text/css"} (gen-statement-css)]])

(defn statement-body [member-name member-order order-date order-total]
  [:body
   [:h1 "Albany Coop"]
   [:h2 (str "Statement for " (member-display-name member-name) " - " order-date)]
   [:table.order
    [:thead
     [:tr
      [:th "Code"]
      [:th "Description"]
      [:th "Packsize"]
      [:th "Case Net"]
      [:th "VAT"]
      [:th "Case Gross"]
      [:th "Albany units/case"]
      [:th "Albany units purchased"]
      [:th "Del?"]
      [:th "Price charged"]
      ]]
     (into [:tbody]
            (for [line member-order]
                    [:tr
                     [:td (:code line)]
                     [:td (:description line)]
                     [:td (:case-size line)]
                     [:td.rightjust (tocurrency (:unit-cost line))]
                     [:td.rightjust (tocurrency (:vat-amount line))]
                     [:td.rightjust (tocurrency (+ (:unit-cost line) (:vat-amount line)))]
                     [:td.rightjust (:albany-units line)]
                     [:td.rightjust (:memdes line)]
                     [:td (:del? line)]
                     [:td.rightjust (tocurrency (:memcost line))]
                     ]
                    ))
    [:tbody
     [:tr
       [:td {:colspan 9} (str "Total for " order-date)]
       [:td.rightjust (tocurrency order-total)]
       ]]
    ]])


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
    (garden.stylesheet/at-media {:print true}
                                [:thead {:display "table-header-group"}])
    )))

(defn order-head [member-name]
  [:head
   [:title (str "Order for " (member-display-name member-name))]
   [:style {:type "text/css"} (gen-order-css)]])

(defn order-body [member-name member-order order-date order-total coordinator]
  [:body
   [:h1 "Albany Food Coop     Order Form"]
   [:table.order
    [:thead
     [:tr
      [:th (str "Name: " (member-display-name member-name))]
      [:th (str "Order Date: " order-date)]
      [:th (str "Coordinator: " coordinator)]
      ]]]
   [:p " "]
   [:table.order
    [:thead
     [:tr
      [:th "Code"]
      [:th "Item"]
      [:th "Unit Q'ty"]
      [:th "Unit Price"]
      [:th "Pref Q'ty"]
      [:th "Actual Q'ty"]
      [:th "Est price"]
      [:th "Final Q'ty"]
      ]]
    (into [:tbody]
          (for [line member-order]
            [:tr
             [:td (:code line)]
             [:td (:description line)]
             [:td (:case-size line)]
             [:td.rightjust (tocurrency (:unit-cost line))]
             [:td.rightjust (:memdes line)]
             [:td " "]
             [:td.rightjust (tocurrency (:memcost line))]
             [:td " "]]
            ))
    [:tbody
     [:tr
      [:td.rightjust {:colspan 6} "Estimated Sub-total for above items "]
      [:td.rightjust (tocurrency order-total)]
      [:td " "]
      ]]

    (into [:tbody]
          (repeat 10
                  [:tr (repeat 8 [:td.space " "])]))
    ]
   ])

(defn emit-statement-html [member-name all-orders order-date]
  (let [member-order (member-name all-orders)
        order-total (reduce #(+ %1 (:memcost %2)) 0 member-order)
        file-name (str (member-display-name member-name) "-" order-date ".html")]
    (println (str "File-name " file-name " Order-total " (format "%.2f" (double order-total)) ))
    (spit  file-name
           (p/html5 (statement-head member-name)
                    (statement-body member-name member-order order-date order-total)))))

(defn emit-order-html [member-name all-orders order-date coordinator]
  (let [member-order (member-name all-orders)
        order-total (reduce #(+ %1 (:memcost %2)) 0 member-order)
        file-name (str (member-display-name member-name) "-" order-date "-OrderForm.html")]
    (println (str "File-name " file-name " Order-total " (format "%.2f" (double order-total)) ))
    (spit  file-name
           (p/html5 (order-head member-name)
                    (order-body member-name member-order order-date order-total coordinator)))))

;; end of html and css-related stuff

(defn generate-statements [spreadsheet-name order-date]
  "Do the work of statement generation."
  (let [wb (load-workbook spreadsheet-name)
        order-sheet (select-sheet "Collated Order" wb)
        all-orders (reduce conj (map #(get-member-order % order-sheet) member-data))]
    (println (str "all-orders keys " (keys all-orders)))
    (println (apply str "all-orders counts " (map (fn [n] (str (name n) ":" (count (n all-orders)) " ")) (keys all-orders))))
    (doall (map #(emit-statement-html % all-orders order-date) (keys member-data)))
    ))

(defn generate-orderforms [spreadsheet-name order-date coordinator]
  "Do the work of order-form generation."
  (let [wb (load-workbook spreadsheet-name)
        sheet (select-sheet "Collated Order" wb)
        all-orders (reduce conj (map #(get-member-order % sheet) member-data))]
    (println (str "all-orders keys " (keys all-orders)))
    (println (apply str "all-orders counts " (map (fn [n] (str (name n) ":" (count (n all-orders)) " ")) (keys all-orders))))
    (doall (map #(emit-order-html % all-orders order-date coordinator) (keys member-data)))
    ))

(defn- usage []
  println "Usage:  -s[tatements] | -o[rder-forms] spreadsheet-name order-date coordinator")

(defn -main
  "Generate all statements from a given Albany spreadsheet."
  [output-type  & args]
   ;; work around dangerous default behaviour in Clojure
  (alter-var-root #'*read-eval* (constantly false))
  (println "output-type is " output-type)
  (case output-type
    "-s"
    (cond (>= (count args) 2)
          (do (println "writing statement files for " (second args))
              (apply generate-statements args))
          :else (usage))
    "-o"
    (cond (>= (count args) 3)
          (do (println "writing order forms for " (second args))
              (apply generate-orderforms args))
          :else (usage))
    (usage)))
