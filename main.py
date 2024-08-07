from flask import Flask, render_template
from app.routes.pccs_tool.route_pccs import pccs_tool

import config

app = Flask(__name__)
app.use_static_for = 'static'
app.config.from_object(config)

@app.route('/app5')
def index():
    return render_template('index.html')

@app.route('/pccs_tool', methods=['GET', 'POST'])
def pccs_tool_route():
    """SCS Tool page"""
    return pccs_tool()

if __name__ == "__main__":
    app.run(debug=True)