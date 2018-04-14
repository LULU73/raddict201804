library(officer)
library(magrittr)

doc <- read_docx("demo/monmodele.docx")

# View(docx_summary(doc))
#
# View(styles_info(x = doc))


doc <- doc %>%
  # suppression du paragraph vide
  body_remove() %>%

  # ajout d'un tableau
  body_add_table(value = head(iris)) %>%

  # nouveau slide et ajout de contenu
  body_add_par(style = "heading 1", value = "district9 est un film cool") %>%
  body_add_img(src = "img/district9-compressed.jpg", height = 409/72/2, width = 801/72/2)

print(doc, target = "demo/result_01.docx")




library(ggplot2)

d <- ggplot(diamonds, aes(carat, price)) +
  geom_hex() + scale_fill_viridis_c() + theme_minimal()

doc <- doc %>%
  body_add_gg(value = d, height = 409/72/2, width = 801/72/2)


print(doc, target = "demo/result_02.docx")


mon_ft <- head(mtcars) %>%
  regulartable() %>%
  color(part = "header", color = "red") %>%
  bold(part = "header", bold = TRUE) %>%
  color(i = ~ mpg < 21, color = "orange") %>%
  set_header_labels(mpg = "Miles per gallon") %>%
  autofit()

doc <- body_add_flextable(doc, value = mon_ft)
print(doc, target = "demo/result_03.docx")



library(mschart)

data <- structure(
  list(
    supp = structure( c(1L, 1L, 1L, 2L, 2L, 2L),
                      .Label = c("OJ", "VC"), class = "factor" ),
    dose = c(0.5, 1, 2, 0.5, 1, 2),
    length = c(13.23, 22.7, 26.06, 7.98, 16.77, 26.14)
  ),
  .Names = c("supp", "dose", "length"),
  class = "data.frame", row.names = c(NA,-6L)
)

monchart <- ms_areachart(data = browser_ts, x = "date",
                      y = "freq", group = "browser")
monchart <- chart_ax_y(monchart, cross_between = "between", num_fmt = "General")
monchart <- chart_ax_x(monchart, cross_between = "midCat", num_fmt = "m/d/yy")



doc <- body_add_chart(doc, chart = monchart, style = "centered")
print(doc, target = "demo/result_04.docx")



