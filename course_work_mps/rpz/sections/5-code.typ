#import "@local/gost732-2017:0.4.2": *


#show: приложение.with(буква: "А", содержание: [ Исходный текст программы ])

#let files = (
  "main.c",
)

#for file in files {
  листинг(raw(read("../code/" + file)))[ Содержимое файла #file ] 
}