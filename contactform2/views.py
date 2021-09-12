from django.shortcuts import render, redirect
from django.core.mail import send_mail, BadHeaderError
from django.http import HttpResponse
from .forms import Form

def contact_form(request):
    if request.method == 'POST':
        form = Form(request.POST)
        if form.is_valid():
            body = {
                'name': form.cleaned_data['name'],
                'email_address': form.cleaned_data['email_address'],
                'subject': form.cleaned_data['subject'],
                'message': form.cleaned_data['message']
            }
            message = '''
                    {}

                    From: {}
                    '''.format(body['message'], body['email_address'])

            send_mail(body['subject'], message, '', ['siddharth.yedlapati@gmail.com'])

            return render(request, 'contactform2/confirmation.html', {})

        return render(request, 'contactform2/index.html', {'form': form})

    else:
        form = Form()
        return render(request, 'contactform2/index.html', {'form': form})