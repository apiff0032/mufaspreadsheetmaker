from django import forms

class TeamForm(forms.Form):
    team_name = forms.CharField(label='MUFA Team Page:', max_length=500)