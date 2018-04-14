library(officer)
library(magrittr)

# chargement d'un doc dans l'objet presentation ----
presentation <- read_pptx("demo/monmodele.pptx")

# Quelques informations ----

View(pptx_summary(presentation))

View(layout_summary(x = presentation))
View(slide_summary(x = presentation))
View(layout_properties(x = presentation))


# Manipulation de contenu ----

presentation <- presentation %>%
  # suppression du slide 4
  remove_slide(index = 4) %>%

  # positionnement sur le slide 3
  on_slide(index = 3) %>%
  # suppression du placeholder 'body'
  ph_remove(type = "body", id_chr = "3") %>%
  # ajout d'un tableau
  ph_with_table(type = "body", index = 1, value = head(iris)) %>%

  # nouveau slide et ajout de contenu
  add_slide(master = "Thème Office", layout = "diapo_std") %>%
  ph_with_text(type = "title", str = "district9 est un film cool") %>%
  ph_with_img(src = "img/district9-compressed.jpg")

print(presentation, target = "demo/result_01.pptx")






# Utilisation de rvg ----

library(ggplot2)
library(rvg)

d <- ggplot(diamonds, aes(carat, price)) +
  geom_hex() + scale_fill_viridis_c() + theme_minimal()

presentation <- add_slide(presentation, master = "Thème Office", layout = "diapo_std")
presentation <- ph_with_gg(presentation, value = d, type = "body")

presentation <- add_slide(presentation, master = "Thème Office", layout = "diapo_std")
presentation <- ph_with_vg(presentation, ggobj = d, type = "body")

print(presentation, target = "demo/result_02.pptx")



# Utilisation de flextable ----
library(flextable)

mon_ft <- head(mtcars) %>%
  regulartable() %>%
  color(part = "header", color = "red") %>%
  bold(part = "header", bold = TRUE) %>%
  color(i = ~ mpg < 21, color = "orange") %>%
  set_header_labels(mpg = "Miles per gallon") %>%
  autofit()
print(mon_ft, preview = "docx")

presentation <- add_slide(presentation, master = "Thème Office", layout = "diapo_std")
presentation <- ph_with_flextable(presentation, value = mon_ft, type = "body")
print(presentation, target = "demo/result_03.pptx")



# Utilisation de mschart ----
library(mschart)

data <- structure(
  list(
    supp = structure( c(1L, 1L, 1L, 2L, 2L, 2L), .Label = c("OJ", "VC"), class = "factor" ),
    dose = c(0.5, 1, 2, 0.5, 1, 2), length = c(13.23, 22.7, 26.06, 7.98, 16.77, 26.14)
  ),
  .Names = c("supp", "dose", "length"),
  class = "data.frame", row.names = c(NA,-6L)
)

monchart <- ms_areachart(data = browser_ts, x = "date",
                      y = "freq", group = "browser")
monchart <- chart_ax_y(monchart, cross_between = "between", num_fmt = "General")
monchart <- chart_ax_x(monchart, cross_between = "midCat", num_fmt = "m/d/yy")



presentation <- add_slide(presentation, master = "Thème Office", layout = "diapo_std")
presentation <- ph_with_chart(presentation, chart = monchart, type = "body")
print(presentation, target = "demo/result_04.pptx")



