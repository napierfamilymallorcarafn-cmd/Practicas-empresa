from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User
from django.conf import settings
import os
import shutil
from django.db.models.signals import post_delete
from django.dispatch import receiver

# Archivo que define las tablas(modelos) de la base de datos de la aplicación 
class Muestra(models.Model):
    # Campos del modelo Muestra
    id_individuo = models.CharField(max_length=20, blank=True, null=True)
    nom_lab = models.CharField(max_length=100, unique=True, help_text="Nombre único de la muestra asignado por el laboratorio")
    id_material = models.CharField(max_length=20,blank=True, null=True, help_text="Material de la muestra")
    volumen_actual = models.FloatField(blank=True, null=True)
    unidad_volumen = models.CharField(max_length=15,blank=True, null=True)
    concentracion_actual = models.FloatField(blank=True, null=True)
    unidad_concentracion = models.CharField(max_length=15,blank=True, null=True)
    masa_actual = models.FloatField(blank=True, null=True)
    unidad_masa = models.CharField(max_length=15,blank=True, null=True)
    fecha_extraccion = models.DateField(blank=True, null=True, help_text="Fecha de extracción de la muestra, en formato AAAA-MM-DD")
    fecha_llegada = models.DateField(blank=True, null=True, help_text="Fecha de llegada de la muestra, en formato AAAA-MM-DD")
    observaciones = models.TextField(blank=True, null=True)
    estado_inicial = models.CharField(max_length=50,blank=True, null=True)
    centro_procedencia = models.CharField(max_length=100,blank=True, null=True)
    lugar_procedencia = models.CharField(max_length=100,blank=True, null=True)
    estado_actual = models.CharField(max_length=50, default='DISP',
                                        choices=[('DISP','Disponible'), ('ENV','Enviada'), ('PENV','Parcialmente enviada'), ('DEST','Destruida')],blank=True, null=True, help_text="Disponible por defecto")
    estudio = models.ForeignKey('Estudio', blank=True, to_field='nombre_estudio', on_delete=models.SET_NULL, null=True)

    # Función para definir un método del modelo muestra que devuelva la posición completa del archivo en el que se situa la misma
    def posicion_completa(self):
        try: 
            sub = self.subposicion
        except Subposicion.DoesNotExist:
            return None
        caja = sub.caja
        posicion_caja_rack = caja.posicion_caja_rack
        rack = caja.rack
        posicion_rack_estante = rack.posicion_rack_estante
        estante = rack.estante
        congelador = estante.congelador
        return(f'{congelador.congelador}-{estante.numero}-{posicion_rack_estante}-{rack.numero}-{posicion_caja_rack}-{caja.numero}-{sub.numero}')
    class Meta:
        # Definición de permisos personalizados para el modelo Muestra
        permissions = [
            ("can_view_muestras_web", "Puede ver muestras en la web"),
            ("can_add_muestras_web", "Puede añadir muestras en la web"),
            ("can_change_muestras_web", "Puede cambiar muestras en la web"),
            ("can_delete_muestras_web", "Puede eliminar muestras en la web"),
        ]
        
    def __str__(self):
        return f"{self.id_individuo} - {self.nom_lab}"

class Localizacion(models.Model):
    # Campos del modelo Localizacion, que referencia a una muestra, usado en la aplicación web en el historial de localizaciones y para definir permisos
    muestra = models.ForeignKey('Muestra',to_field = "nom_lab",related_name="localizacion", blank=True, null=True, on_delete=models.SET_NULL)
    congelador = models.CharField(max_length=50, blank=True, null=True)
    estante = models.CharField(max_length=50,blank=True, null=True)
    rack = models.CharField(max_length=50,blank=True, null=True)
    caja = models.CharField(max_length=50,blank=True, null=True)
    subposicion = models.CharField(max_length=50,blank=True, null=True)

    class Meta:
        # Permisos asociados al modelo localización
        permissions = [
            ("can_view_localizaciones_web", "Puede ver localizaciones en la web"),
            ("can_add_localizaciones_web", "Puede añadir localizaciones en la web"),
            ("can_change_localizaciones_web", "Puede cambiar localizaciones en la web"),
            ("can_delete_localizaciones_web", "Puede eliminar localizaciones en la web"),
        ]
    def __str__(self):
        return f"{self.congelador} - {self.estante} - {self.rack} - {self.caja} - {self.subposicion}"


class Congelador(models.Model):
    # Modelo congelador, que lista los datos y el nombre de los congeladores disponibles 
    congelador = models.CharField(max_length=50, unique=True)
    modelo = models.CharField(max_length=50,blank=True, null=True)
    temperatura_minima = models.CharField(max_length=50,blank=True, null=True)
    temperatura_maxima = models.CharField(max_length=50,blank=True, null=True)
    localizacion_edificio = models.CharField(max_length=50,blank=True, null=True)
    fotografia = models.ImageField(upload_to='congeladores/', blank=True, null=True)

    def save(self, *args, **kwargs):
        if self.pk:
            old = Congelador.objects.get(pk=self.pk)
            if old.congelador != self.congelador:
                Localizacion.objects.filter(congelador=old.congelador).update(congelador=self.congelador)
        super().save(*args, **kwargs)


class Estante(models.Model):
    # Modelo estante, que está relacionado con un congelador
    congelador = models.ForeignKey(Congelador, on_delete=models.CASCADE, related_name='estantes', to_field='congelador')
    numero = models.CharField(max_length=50)
    class Meta:
        # Un estante en concreto solo puede estar asociado a un congelador
        constraints = [
            models.UniqueConstraint(
                fields=['congelador', 'numero'],
                name='unique_estante_por_congelador'
            )
        ]

    def save(self, *args, **kwargs):
        if self.pk:
            old = Estante.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(congelador=self.congelador.congelador, estante=old.numero).update(estante=self.numero)
        super().save(*args, **kwargs)


class Rack(models.Model):
    # Modelo rack, que está relacionado con un estante y contiene la posición del mismo dentro del estante
    estante = models.ForeignKey(Estante, on_delete=models.CASCADE, related_name='racks')
    numero = models.CharField(max_length=50)
    posicion_rack_estante = models.CharField(max_length=50)
    class Meta:
        # Un rack en concreto solo puede estar asociado a un estante
        constraints = [
            models.UniqueConstraint(
                fields=['estante', 'numero'],
                name='unique_rack_por_estante'
            ),
            models.UniqueConstraint(
                fields=['estante', 'posicion_rack_estante'],
                name='unique_posicion_rack_por_estante'
            )
        ]

    def save(self, *args, **kwargs):
        if self.pk:
            old = Rack.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(
                    congelador=self.estante.congelador.congelador,
                    estante=self.estante.numero,
                    rack=old.numero
                ).update(rack=self.numero)
        super().save(*args, **kwargs)


class Caja(models.Model):
    # Modelo caja, que está relacionado con un rack y contiene la posición de la misma dentro del rack
    rack = models.ForeignKey(Rack, on_delete=models.CASCADE, related_name='cajas')
    numero = models.CharField(max_length=50)
    posicion_caja_rack = models.CharField(max_length=50)
    class Meta:
        # Una caja en concreto solo puede estar asociada a un rack
        constraints = [
            models.UniqueConstraint(
                fields=['rack', 'numero'],
                name='unique_caja_por_rack'
            ),
            models.UniqueConstraint(
                fields=['rack', 'posicion_caja_rack'],
                name='unique_posicion_caja_por_rack'
            )
        ]

    def save(self, *args, **kwargs):
        if self.pk:
            old = Caja.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(
                    congelador=self.rack.estante.congelador.congelador,
                    estante=self.rack.estante.numero,
                    rack=self.rack.numero,
                    caja=old.numero
                ).update(caja=self.numero)
        super().save(*args, **kwargs)


class Subposicion(models.Model):
    # Modelo subposicion, que está relacionado con un congelador
    caja = models.ForeignKey(Caja, on_delete=models.CASCADE, related_name='subposiciones')
    numero = models.CharField(max_length=50)
    vacia = models.BooleanField(default=True)
    muestra = models.OneToOneField(Muestra, on_delete=models.SET_NULL, null=True, blank=True, related_name='subposicion')
    class Meta:
        # Una subposicion en concreto solo puede estar asociada a una caja
        constraints = [
            models.UniqueConstraint(
                fields=['caja', 'numero'],
                name='unique_subposicion_por_caja'
            )
        ]

    def save(self, *args, **kwargs):
        if self.pk:
            old = Subposicion.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(
                    congelador=self.caja.rack.estante.congelador.congelador,
                    estante=self.caja.rack.estante.numero,
                    rack=self.caja.rack.numero,
                    caja=self.caja.numero,
                    subposicion=old.numero
                ).update(subposicion=self.numero)
        super().save(*args, **kwargs)

class historial_localizaciones(models.Model):
    # Modelo que guarda el historial de localizaciones de una muestra
    muestra = models.ForeignKey('Muestra',related_name="historial_localizaciones",on_delete=models.CASCADE)
    localizacion = models.ForeignKey('Localizacion',related_name="historial_localizaciones",on_delete=models.SET_NULL, null=True, blank=True)
    fecha_asignacion = models.DateField(default=timezone.now) 
    usuario_asignacion = models.ForeignKey(User,on_delete=models.PROTECT, blank=True, null=True)   

class registro_destruido(models.Model):
    # Modelo que guarda el registro del estado 'destruida' de una muestra
    muestra = models.ForeignKey('Muestra',related_name="estado_destruido",on_delete=models.CASCADE)
    fecha = models.DateField(default = timezone.now)
    usuario = models.ForeignKey(User,on_delete=models.PROTECT, blank=True, null=True)
class Estudio(models.Model):
    # Campos del modelo Estudio
    referencia_estudio = models.CharField(max_length=100, blank=True, null=True)
    nombre_estudio = models.CharField(max_length=100,unique=True)
    descripcion_estudio = models.TextField(blank=True, null=True)
    fecha_inicio_estudio = models.DateField(blank=True, null=True)
    fecha_fin_estudio = models.DateField(blank=True, null=True)
    investigador_principal = models.CharField(max_length=100, blank=True, null=True)
    investigadores_asociados = models.ManyToManyField(settings.AUTH_USER_MODEL, limit_choices_to={'groups__name': 'Investigadores'},related_name="estudios_asignados",blank = True)
    class Meta:
        # Permisos asociados al modelo estudio
        permissions = [
            ("can_view_estudios_web", "Puede ver estudios en la web"),
            ("can_add_estudios_web", "Puede añadir estudios en la web"),
            ("can_change_estudios_web", "Puede cambiar estudios en la web"),
            ("can_delete_estudios_web", "Puede eliminar estudios en la web"),
        ]
    def __str__(self):
        return f"Estudio {self.nombre_estudio}"
class historial_estudios(models.Model):
    # Modelo que guarda el historial de estudios de una muestra
    muestra = models.ForeignKey('Muestra',related_name="historial_estudios",on_delete=models.CASCADE)
    estudio = models.ForeignKey('Estudio',related_name="historial_estudios",on_delete=models.SET_NULL, blank=True, null=True)
    fecha_asignacion = models.DateField(default=timezone.now)
    usuario_asignacion = models.ForeignKey(User,on_delete=models.PROTECT,blank=True, null=True) 

def ruta_documentos(instance,filename):
    # Función que define la carpeta asociada a cada estudio para guardar los documentos 
    return f"estudios/{instance.estudio.id}/{filename}"
class Documento(models.Model):
    # Campos del modelo documento, asociado a un estudio
    estudio = models.ForeignKey('Estudio',related_name = "estudio", on_delete=models.CASCADE)
    archivo = models.FileField(upload_to=ruta_documentos)
    fecha_subida = models.DateTimeField(auto_now_add=True)
    categoria = models.CharField(blank=True, null=True, max_length=50)
    usuario_subida = models.ForeignKey(User,on_delete=models.PROTECT)
    descripcion = models.TextField(blank=True, null=True)
    eliminado = models.BooleanField(default = False)
    fecha_eliminacion = models.DateField(blank = True, null=True)
class Envio(models.Model):
    # Campos del modelo Envio, que referencia a una muestra
    muestra = models.ForeignKey('Muestra',related_name="envio",on_delete=models.CASCADE)
    volumen_enviado = models.FloatField()
    unidad_volumen_enviado = models.CharField(max_length=15)
    concentracion_enviada = models.FloatField()
    unidad_concentracion_enviada = models.CharField(max_length=15)
    centro_destino = models.CharField(max_length=100)
    lugar_destino = models.CharField(max_length=100)
    fecha_envio = models.DateField(default=timezone.now)
    usuario_envio = models.ForeignKey(User, on_delete=models.PROTECT, blank=True, null=True)

    def __str__(self):
        return f"Envio de Muestra {self.id_individuo} - {self.nom_lab} el {self.fecha_envio}"
    
class agenda_envio(models.Model):
    # Campos de los datos de contacto de los centros de envio
    centro = models.CharField(max_length=200,unique=True,default=None)
    lugar=models.CharField(max_length=200)
    direccion=models.TextField()
    persona_contacto = models.CharField(max_length=200,blank=True, null=True)
    telefono_contacto=models.IntegerField(blank=True, null=True)


@receiver(post_delete, sender=Documento)
def eliminar_archivo_documento(sender, instance, **kwargs):
    """Eliminar el archivo físico cuando se borra la fila Documento."""
    try:
        if instance.archivo:
            instance.archivo.delete(save=False)
    except Exception:
        # No interrumpir la eliminación por errores en el borrado del fichero
        pass


@receiver(post_delete, sender=Estudio)
def eliminar_carpeta_estudio(sender, instance, **kwargs):
    """Eliminar la carpeta `media/estudios/<id>` tras borrar el estudio (si existe)."""
    try:
        media_root = settings.MEDIA_ROOT
        carpeta = os.path.join(media_root, 'estudios', str(instance.id))
        if os.path.exists(carpeta) and os.path.isdir(carpeta):
            shutil.rmtree(carpeta)
    except Exception:
        # Silenciar errores de borrado de carpeta para no bloquear la eliminación del objeto
        pass