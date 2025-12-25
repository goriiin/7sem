
#import "@local/gost732-2017:0.4.2": *
#import "@local/bmstu:0.3.0": *

#show: приложение.with(буква: "В", содержание: [ Перечень элементов ])


#страница(
   image("/assets/spisok1.jpg", width: 100%, fit: "cover"),
   повернуто: нет,
   формат: "a4",
  номер: нет
)
#страница(
   image("/assets/spisok2.jpg", width: 100%, fit: "cover"),
   повернуто: нет,
   формат: "a4",
  номер: нет
)