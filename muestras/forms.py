from django import forms
from .models import Muestra, Localizacion, Estudio, Documento,agenda_envio, Congelador

# Archivo que describe los formularios de la app basados en modelos 
class MuestraForm(forms.ModelForm):
    # Formulario basado en el modelo Muestra, se incluyen todos los campos del modelo
    class Meta:
        model = Muestra
        fields = '__all__'
        widgets = {
            'fecha_extraccion': forms.DateInput(format='%Y-%m-%d', attrs={'type': 'date'}),
            'fecha_llegada': forms.DateInput(format='%Y-%m-%d', attrs={'type': 'date'}),
        }

class UploadExcel(forms.Form):
    # Formulario para subir un archivo Excel 
    excel_file = forms.FileField(required=True)

class archivar_muestra_form(forms.ModelForm):
        # Formulario para archivar una muestra en una localización específica
        class Meta:
            model = Localizacion
            exclude = ('muestra', )
class EstudioForm(forms.ModelForm):
    # Formulario para añadir un estudio, se incluyen algunos campos
    class Meta:
        model=Estudio
        fields= ['referencia_estudio','nombre_estudio','descripcion_estudio','fecha_inicio_estudio','fecha_fin_estudio','investigador_principal']
        widgets = {
            'fecha_inicio_estudio': forms.DateInput(format='%Y-%m-%d', attrs={'type': 'date'}),
            'fecha_fin_estudio': forms.DateInput(format='%Y-%m-%d', attrs={'type': 'date'}),
        }
class DocumentoForm(forms.ModelForm):
    # Formulario para añadir un documento, se incluyen algunos campos
    class Meta:
        model = Documento 
        fields = ['archivo','categoria','descripcion']

class Centroform(forms.ModelForm):
    # Formulario para añadir un nuevo centro de envio a la agenda de envios
    class Meta:
        model=agenda_envio
        fields = '__all__'

class Congeladorform(forms.ModelForm):
    # Formulario para añadir un congelador nuevo 
    class Meta:
        model=Congelador
        fields = '__all__'