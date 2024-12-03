(ns albany-statements.output-pdf
  (:require [clojure.string :as string]
            [hiccup
             [core :as h]
             [page :as p]]
            [garden.core]
            [garden.stylesheet]
            [clj-htmltopdf.core :as pdf]
            [albany-statements
             [balance :as bal]
             [cli :refer [version-tostring]]
             [util :as u]]))

(defn member-display-name
  [name-keyword]
  (string/capitalize (name name-keyword)))

;; start of html and css-realted functions
(def dark-blue "#A4D1FE")
(def light-blue "#EAF4FF")
(def light-grey "#ECECEA")

(def pdf-options {:page {:margin      :narrow
                         :size        :A4
                         :orientation :landscape
                         :margin-box  {:bottom-right {:paging ["Page " :page " of " :pages]}} }
                  :styles {:font-size "10pt"
                            :color     "#000"}
                  :debug {:display-html? false
                          :display-options? false}})

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
                      revision
                      suffix]
  (let [suffix-string (if suffix (str " (" suffix ")") "")]
    [:body
     [:h1 "Albany Coop"]
     [:h2 (str "Statement for "
               (member-display-name member-name)
               " - "
               order-date
               suffix-string)]
     [:div.order
      [:table.order
       [:thead
        [:tr
         [:th {:rowspan 2} "Code"]
         [:th {:rowspan 2} "Description"]
         [:th {:rowspan 2} "Case"]
         [:th {:colspan 3} "Price (single)"]
         [:th {:rowspan 2} "Singles per case"]
         [:th {:colspan 2} "Amount purchased"]
         [:th {:rowspan 2}"Status"]
         [:th {:rowspan 2}"Price charged"]]
        [:tr
         [:th "Net"]
         [:th "VAT"]
         [:th "Gross"]
         [:th "Singles"]
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
                [:td.rightjust (:singles-per-case line)]
                [:td.rightjust (:memdes line)]
                [:td.rightjust (u/essential-cases (:memdes line) (:singles-per-case line))]
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
     [:div#save-revision
      [:span (format "Taken from %s. Save revision %d" spreadsheet-name (int revision))]]]))


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

(defn draft-order-table-head []
  [:thead
   [:tr
    [:th {:rowspan 2} "Code"]
    [:th {:rowspan 2} "Item"]
    [:th {:rowspan 2} "Case"]
    [:th {:rowspan 2} "Price (single)"]
    [:th {:rowspan 2} "Unit VAT"]
    [:th {:rowspan 2} "Singles per case"]
    [:th {:colspan 2} "Pref Q'ty"]
    [:th {:colspan 2} "Actual Q'ty"]
    [:th {:rowspan 2} "Est total price"]
    [:th {:rowspan 2} "Final Q'ty"]]
   [:tr
    [:th "Singles"]
    [:th "Cases"]
    [:th "Singles"]
    [:th "Cases"]]])

(defn draft-order-table-row [line]
  [:tr
   [:td [:b (:code line)]]
   [:td (u/format-description (:description line))]
   [:td (:case-size line)]
   [:td.rightjust (u/tocurrency (:unit-cost line))]
   [:td.rightjust (u/tocurrency (:vat-amount line))]
   [:td.rightjust (:singles-per-case line)]
   [:td.rightjust (:memdes line)]
   [:td.rightjust (u/essential-cases (:memdes line) (:singles-per-case line))]
   [:td " "]
   [:td " "]
   [:td.rightjust (u/tocurrency (:memcost line))]
   [:td " "]])

(defn final-order-table-head []
  [:thead
   [:tr
    [:th {:rowspan 2} "Code"]
    [:th {:rowspan 2} "Item"]
    [:th {:rowspan 2} "Case"]
    [:th {:rowspan 2} "Unit Price"]
    [:th {:rowspan 2} "Unit VAT"]
    [:th {:rowspan 2} "Est total price"]
    [:th {:rowspan 2} "Singles per case"]
    [:th {:colspan 2} "Quantity Ordered"]
    [:th {:rowspan 2} "Quantity Delivered"]]
   [:tr
    [:th "Singles"]
    [:th "Cases"]]])

(defn final-order-table-row [line]
  [:tr
   [:td [:b (:code line)]]
   [:td (u/format-description (:description line))]
   [:td (:case-size line)]
   [:td.rightjust (u/tocurrency (:unit-cost line))]
   [:td.rightjust (u/tocurrency (:vat-amount line))]
   [:td.rightjust (u/tocurrency (:memcost line))]
   [:td.rightjust (:singles-per-case line)]
   [:td.rightjust (:memdes line)]
   [:td.rightjust (u/essential-cases (:memdes line) (:singles-per-case line))]
   [:td.rightjust (if-not (nil? (:del? line)) (str "(" (:del? line) ")") " ")]])

(defn order-body [member-name
                  member-order
                  balance
                  order-date
                  order-total
                  coordinator
                  version
                  suffix]
  (let [is-draft    (= version :d)
        num-cols-before    (case version :d 10 :f 5)
        num-cols-after (case version :d 1 :f 4)
        blank-lines (case version :d 5 :f 4)
        suffix-string (if suffix (str " (" suffix ")") "")
        [_ balance-brought-forward] (bal/bal-item :bf balance)
        who-owes-who (cond
                       (< balance-brought-forward 0) " (you owe Albany)"
                       (> balance-brought-forward 0) " (Albany owes you)"
                       :else "")]
    [:div
     [:h1 "Albany Food Coop     Order Form"]
     (case version
       :d [:h2 (str "Pre-meeting DRAFT" suffix-string)]
       :f [:h2 (str "FINAL for order sorting" suffix-string)])
     [:table.order
      [:thead
       [:tr
        [:th [:span "Name: " [:b.bigtext (member-display-name member-name)]]]
        [:th [:span "Order Date: " [:b.bigtext order-date]]]
        [:th (str "Coordinator: " coordinator)]
        ]]]
     [:p " "]
     [:table.order
      (case version
        :d (draft-order-table-head)
        :f (final-order-table-head))
      (into [:tbody]
            (for [line member-order]
              (case version
                :d (draft-order-table-row line)
                :f (final-order-table-row line))))
      [:tbody
       [:tr
        [:td.rightjust {:colspan num-cols-before} [:b "Estimated Sub-total for above items "]]
        [:td.rightjust [:b (u/tocurrency order-total)]]
        [:td {:colspan num-cols-after} " "]
        ]
       (when is-draft 
         [:tr
          [:td.rightjust {:colspan num-cols-before} [:b (str "Balance from previous order" who-owes-who)]]
          [:td.rightjust [:b (u/tocurrency balance-brought-forward)]]
          [:td {:colspan num-cols-after} " "]
          ])
       (when is-draft
         [:tr
          [:td.rightjust {:colspan num-cols-before} [:b "Estimated amount to pay" ]]
          [:td.rightjust [:b (u/tocurrency (- order-total balance-brought-forward))]]
          [:td {:colspan num-cols-after} " "]
          ])]

      (into [:tbody]
            (repeat blank-lines
                    [:tr (repeat (+ 1 num-cols-before num-cols-after) [:td.space " "])]))
      ]
     ]))

(defn emit-statement-html [member-name
                           all-orders
                           mem-balance
                           order-date
                           spreadsheet-name
                           revision
                           suffix]
  (let [member-order (member-name all-orders)
        order-total (reduce #(+ %1 (:memcost %2)) 0 member-order)
        fname-suffix (if suffix (str "-" suffix) "")
        person-name (member-display-name member-name)
        footer-text (str "Statement for " person-name)
        file-name (str person-name
                       "-"
                       order-date
                       fname-suffix
                       ".pdf")
        options        (update-in pdf-options
                                  [:page :margin-box]
                                  merge
                                  {:bottom-left {:element "save-revision"}
                                   :bottom-center {:text footer-text}})
        doc-options    {:doc {:title  footer-text
                              :author "Albany Coop"
                              :subject order-date}}]
    (println (str "File-name " file-name " Order-total " (format "%.2f" (double order-total)) ))
    (pdf/->pdf
     (p/html5 (statement-head member-name)
              (statement-body member-name
                              member-order
                              mem-balance
                              order-date
                              order-total
                              spreadsheet-name
                              revision
                              suffix))
     file-name
     (merge options doc-options))))

(defn emit-order-html [member-name
                       all-orders
                       mem-balance
                       order-date
                       coordinator
                       version
                       suffix]
  (let [member-order (member-name all-orders)
        order-total (reduce #(+ %1 (:memcost %2)) 0 member-order)
        fname-suffix (if suffix (str "-" suffix) "")
        person-name (member-display-name member-name)
        display-version (string/upper-case (version-tostring version))
        footer-text (str person-name " " display-version " order form")
        file-name (str person-name
                       "-"
                       order-date
                       "-OrderForm-"
                       display-version
                       fname-suffix
                       ".pdf")
        options        (update-in pdf-options
                                  [:page :margin-box]
                                  merge
                                  {:bottom-left {:text footer-text}})
        doc-options    {:doc {:title  footer-text
                              :author "Albany Coop"
                              :subject order-date}}]
    (println (str "File-name " file-name " Order-total " (format "%.2f" (double order-total)) ))

    (pdf/->pdf
     (p/html5 (order-head member-name)
              (order-body member-name
                          member-order
                          mem-balance
                          order-date
                          order-total
                          coordinator
                          version
                          suffix)) 
     file-name
     (merge options doc-options))))

;; end of html and css-related stuff
