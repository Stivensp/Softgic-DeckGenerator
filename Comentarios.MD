tuve que hacer un editor visual para poder tener un control total del diseño
ya que al trabajar con algo con diseño no se puede hacer todo por cli ni
mover los elementos con variables o algo asi, el mejor era powerpoint como
la herramienta de diseño nada mejor que esa herramienta

Las medidas standard de la presentacion cuando agarra el diseño no eran las que eran quedaba mas pequeño del standard al crear el deck

Se lee muy bien el yaml y lo incorpora bien en el diseño pero... hace falta 
el diseño real de como se espera el deck realmente control en cuanto al diseño a la generacion del diseño

En cuanto a las imagenes me estaban sugiriendo que suba las imagenes directamente a una carpeta y luego escriba sus coordenadas en el slide pero esto tomaria demasiado tiempo entonces el me dice que para que yo suba la imagen al config-fvisual que seria el template guia de todo el programa y que toma el mapa de bits y luego lo sube directamente a repositorio con sus coordenadas

Las formas y la edicion en el diseño al principio daban errores, si eran circulos o triangulos los pasaba a cuadrados, tambien se mantenian formas y assets que no deberian de salir, se solucionó

El corazón del programa es el lector del config-visual ya que copia las coordenadas de cada una de las figuras o textos y el lector yml pone la informacion que se desea para poder automatizar

Cuando borraba un elemento del powerpoint y volvia a abrir el editor visual el elemento volvia a aparecer como si nada, eso fue un problema de diseño bastante molesto porque uno lo borra y sigue saliendo en el deck final tambien, se tuvo que agregar un sistema donde si el elemento no esta en el powerpoint queda marcado como eliminado y el programa lo respeta en todos lados

Habian formas que powerpoint tiene pero la libreria de python no las reconoce, cosas de versiones nuevas de office que no estan en el catalogo, entonces en esos casos toca leer el xml directamente del archivo para saber que tipo de forma es y reproducirla de esa manera, es una solucion que funciona aunque no es lo mas bonito del mundo

hubo un error muy silencioso que era dificil de encontrar, cuando el script corre desde la carpeta tools python no agrega el root del proyecto al path sino la carpeta del script entonces todas las importaciones del proyecto fallaban sin dar ningun error visible simplemente no cargaba nada y el programa seguia corriendo como si todo estuviera bien pero sin datos, se tuvo que forzar el path manualmente al inicio de cada script

La parte de imagenes al principio era un dolor de cabeza porque la unica forma que se proponia era escribir coordenadas a mano en el yaml lo cual es absurdo si ya tienes powerpoint, entonces se diseño de manera que uno inserta la imagen directamente en el editor visual de powerpoint la posiciona como quiere y cuando aplica los cambios el programa extrae la imagen del archivo y la guarda en el repositorio con todas sus coordenadas automaticamente sin tocar nada de yaml

las imagenes de fondo son un caso especial porque tienen que quedar detras de todos los demas elementos del slide, si se agregan normal quedan encima de todo entonces se tuvo que hacer que si la imagen cubre mas de la mitad del slide se mueva automaticamente al fondo de la pila de capas del powerpoint