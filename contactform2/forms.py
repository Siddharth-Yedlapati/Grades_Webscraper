from django import forms

class Form(forms.Form):
    name = forms.CharField(max_length = 100)
    email_address = forms.EmailField(max_length = 200)
    subject = forms.CharField(max_length = 100)
    message = forms.CharField(widget = forms.Textarea, max_length = 1000)
