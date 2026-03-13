from django import forms
from django.core.exceptions import ValidationError
from .models import Muestra, Localizacion, Estudio, Documento,agenda_envio, Congelador

# Validador que rechaza el carácter punto y coma (;)
def no_semicolon(value):
    """Valida que el valor no contenga punto y coma (;)"""
    if value and ';' in str(value):
        raise ValidationError('El carácter ";" no está permitido en este campo.')

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
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # ESTADO ACTUAL ES OBLIGATORIO
        if 'estado_actual' in self.fields:
            self.fields['estado_actual'].required = True
        
        # Agregar validador a campos de texto que no deben contener ;
        text_fields = ['id_individuo', 'nom_lab', 'id_material', 'centro_procedencia', 
                      'lugar_procedencia', 'estado_actual', 'observaciones', 'estado_inicial']
        for field_name in text_fields:
            if field_name in self.fields:
                if not isinstance(self.fields[field_name].validators, list):
                    self.fields[field_name].validators = list(self.fields[field_name].validators)
                self.fields[field_name].validators.append(no_semicolon)

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
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Agregar validador a campos de texto que no deben contener ;
        text_fields = ['referencia_estudio', 'nombre_estudio', 'descripcion_estudio', 'investigador_principal']
        for field_name in text_fields:
            if field_name in self.fields:
                if not isinstance(self.fields[field_name].validators, list):
                    self.fields[field_name].validators = list(self.fields[field_name].validators)
                self.fields[field_name].validators.append(no_semicolon)

class DocumentoForm(forms.ModelForm):
    # Formulario para añadir un documento, se incluyen algunos campos
    class Meta:
        model = Documento 
        fields = ['archivo','categoria','descripcion']
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Agregar validador a campos de texto que no deben contener ;
        text_fields = ['categoria', 'descripcion']
        for field_name in text_fields:
            if field_name in self.fields:
                if not isinstance(self.fields[field_name].validators, list):
                    self.fields[field_name].validators = list(self.fields[field_name].validators)
                self.fields[field_name].validators.append(no_semicolon)

class Centroform(forms.ModelForm):
    # Formulario para añadir un nuevo centro de envio a la agenda de envios
    class Meta:
        model=agenda_envio
        fields = '__all__'
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Agregar validador a todos los campos de texto para no contener ;
        for field_name, field in self.fields.items():
            if isinstance(field, forms.CharField):
                if not isinstance(field.validators, list):
                    field.validators = list(field.validators)
                field.validators.append(no_semicolon)

class Congeladorform(forms.ModelForm):
    # Formulario para añadir un congelador nuevo 
    class Meta:
        model=Congelador
        fields = '__all__'
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Agregar validador a todos los campos de texto para no contener ;
        for field_name, field in self.fields.items():
            if isinstance(field, forms.CharField):
                if not isinstance(field.validators, list):
                    field.validators = list(field.validators)
                field.validators.append(no_semicolon)