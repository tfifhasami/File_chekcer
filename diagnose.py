import sys
import os

print("Testing template rendering...")
print("=" * 60)

# Set up paths
application_path = os.path.dirname(os.path.abspath(__file__))
template_folder = os.path.join(application_path, 'templates')

print(f"Application path: {application_path}")
print(f"Template folder: {template_folder}")
print(f"Template exists: {os.path.exists(template_folder)}")

if os.path.exists(template_folder):
    print(f"Files in template folder: {os.listdir(template_folder)}")
    
    template_file = os.path.join(template_folder, 'index.html')
    print(f"index.html exists: {os.path.exists(template_file)}")
    
    if os.path.exists(template_file):
        print(f"index.html size: {os.path.getsize(template_file)} bytes")
        
        # Try to read the file
        try:
            with open(template_file, 'r', encoding='utf-8') as f:
                content = f.read()
                print(f"✓ Successfully read index.html ({len(content)} characters)")
        except Exception as e:
            print(f"✗ Error reading index.html: {e}")

print("\n" + "=" * 60)
print("Testing Flask...")

try:
    from flask import Flask, render_template
    print("✓ Flask imported")
    
    app = Flask(__name__, template_folder=template_folder)
    print("✓ Flask app created")
    
    # Test rendering outside of request context
    with app.app_context():
        try:
            # This won't work but will give us the error
            from flask import render_template_string
            test_html = "<h1>Test</h1>"
            result = render_template_string(test_html)
            print("✓ render_template_string works")
        except Exception as e:
            print(f"✗ render_template_string error: {e}")
    
    # Now test with actual server
    @app.route('/')
    def index():
        try:
            print("\n>>> Entering index route")
            result = render_template('index.html')
            print(">>> Template rendered successfully")
            return result
        except Exception as e:
            print(f">>> ERROR in route: {e}")
            import traceback
            traceback.print_exc()
            return f"<pre>{traceback.format_exc()}</pre>", 500
    
    print("\nStarting Flask server on http://127.0.0.1:5555")
    print("Open this URL in your browser")
    print("=" * 60 + "\n")
    
    # Run without reloader to see errors
    app.run(host='127.0.0.1', port=5555, debug=False, use_reloader=False)
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()