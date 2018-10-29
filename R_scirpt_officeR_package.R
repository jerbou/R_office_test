# ====== officer test ======


# ==== 00 : chargement de librairie ====
library(dplyr)
library(XML)
library(plyr)
library(xmlparsedata)
library(data.table)
library(dtplyr)
library(ggplot2)
library(rjson)
library(XLConnect)
library(xlsimple)
library(RODBC)
library(xlsx)
library(data.table)
library(readr)
library(sqldf)
library(stringr)
library(tidyr)
library(readODS)
library(dismo)
library(XML)
library(stringi)
library(httr)
library(doBy)
library(scales)
library(banR)
library(readr)
library(readxl)
library(gridExtra)
library(forcats)
library(officer)
library(mschart)
library(rvg)

# chargement du workdir
# setwd('T:/Géobourgogne/DONNEES/PLATEFORME/ENTRANTE/ternum/')
setwd('C:/COPY_data_local/ternum/')

# df0 <- read_excel("T:/Géobourgogne/DONNEES/PLATEFORME/ENTRANTE/ternum/Liste adhérents pour météo_03_2018.xlsx")
# df0 <- read_excel("C:/COPY_data_local/ternum/Liste adhérents pour météo_03_2018.xlsx")
df0 <- read_excel("C:/COPY_data_local/ternum/Liste adhérents pour météo_04_2018.xlsx")
# === Reverse order tips ====
# https://stackoverflow.com/questions/42710056/reverse-stacked-bar-order




# https://davidgohel.github.io/officer/articles/offcran/graphics.html



# graphe
h0 <- ggplot(data=df0, aes(x = factor(df0$collège))) +geom_bar()
h1 <- h0 + labs(fill= "Type local", x="college", y = "nombre") + geom_text(color = "white", stat='count', aes(label=..count..), vjust=1.25)

# table
resume <- df0 %>%
  group_by(nature_ad) %>%
  summarize_each(funs(length, min, median, mean, max, sum), cotisation_annuelle)
names(resume) <- c("college","nombre","min","medianne","moyen","max", "somme")
resume$moyen <- round(resume$moyen,0)
resume$somme <- round(resume$somme,0)
resume2 <- resume[order(-resume$nombre),]
knitr::kable(resume2)

# produce an emf file containing the ggplot
#filename <- tempfile(fileext = ".emf")
# emf(file = "filename.emf", width = 6, height = 7)
# print(h1)
# dev.off()


# my_barchart <- ms_barchart(data = df0,
#                           x = "collège", y = "cotisation_annuelle", group = "CodeDepartement")
# my_barchart <- chart_settings( x = my_barchart,
#                                dir="vertical", grouping="clustered", gap_width = 50 )

# read_pptx() %>% 
#   add_slide(layout = "Title and Content", master = "Office Theme") %>% 
#   ph_with_chart(chart = my_barchart , type = "body") %>% 
#   add_slide(layout = "Title and Content", master = "Office Theme") %>% 
#   ph_with_chart_at(chart = my_barchart , 
#                    left = 0, top = 0, width = 4, height = 4) %>% 
#   print(target = "demo_mschart_01.pptx")
# 
# # creation d'un ppt vierge
# read_pptx() %>% 
#   add_slide(layout = "Title and Content", master = "Office Theme") %>% 
#   ph_with_vg(code = print(h1), type = "body") %>% 
#   print(target = "demo_rvg.pptx")
# # Wuhuuuuuuuuuuuuuuu !!

# ==== ADD a new slide to an existing powerpoint ====
# on lit un ppt 
my_pres <- read_pptx("test_pres.pptx")
# add slide
# https://davidgohel.github.io/officer/articles/powerpoint.html

layout_summary(my_pres)
slide_summary(my_pres)

my_pres <- my_pres %>%  
  add_slide(layout = "Vide", master = "ModelePrésentationGIP2015") %>% 
  # ph_empty(type= "body") %>%
  ph_with_text(type = "title", str = "Nombre d'adhérents par collège") %>%
  ph_with_vg_at(code = print(h1), height = 6, width = 8, left = 1, top = 1) %>% 
# utilisation de la fonction custom ph_with_vg_ad avec parametre
  add_slide(layout = "Vide", master = "ModelePrésentationGIP2015") %>%
  ph_with_table_at(value = resume2, height =4 , width = 8, left = 0.5, top = 0.5, first_row = TRUE ) %>%
  print(my_pres, target = "demo_rvg_add.pptx")

layout_properties ( x = my_pres, layout = "Two Content", master = "Office Theme" ) %>% head()

annotate_base(output_file = "annotated_layout.pptx")
