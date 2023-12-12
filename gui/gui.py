from formation import AppBuilder


def event_promo(event=None):
    # event parameter needs to be there because using the bind method passes an event object
    # access the expr_var we created earlier to determine the current expression entered
    # filepath = app..get()
    # display the result
    print("event promo")

app = AppBuilder(path="gui.xml")


app.connect_callbacks(globals())

app.mainloop()
