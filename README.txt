PASOS PARA DESPLEGAR 'Riesgos y AST' EN RENDER

1. Entra a https://render.com y crea una cuenta (si no tienes una).
2. Crea un nuevo proyecto (New Web Service).
3. Conecta tu GitHub o sube manualmente este ZIP como nuevo repositorio.
4. En configuración:
   - Runtime: Python 3
   - Start command: python main.py
   - Build command: pip install -r requirements.txt
5. En Environment:
   - Web service
   - Public
   - Free instance

6. Cuando se despliegue, Render te dará una URL como:
   https://riesgos-ast.onrender.com

7. Usa esa URL en tu esquema OpenAPI:
   servers:
     - url: https://riesgos-ast.onrender.com

¡Listo!