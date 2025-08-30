# Installing Export for My Efficient Elements to Applications Folder

## 🚀 Quick Installation

### Method 1: Drag and Drop (Recommended)
1. **Open Finder** and navigate to your project folder: `/Users/upmanyuskand/repos/ee_my_elements`
2. **Open Applications folder** in a new Finder window (`Cmd+Shift+A`)
3. **Drag** `Launch EE GUI.app` from the project folder to the Applications folder
4. **Launch** from Applications folder, Spotlight, or Dock

### Method 2: Terminal Copy
```bash
# Copy the app to Applications folder
cp -R "/Users/upmanyuskand/repos/ee_my_elements/Launch EE GUI.app" /Applications/

# Launch from Applications
open "/Applications/Launch EE GUI.app"
```

## ✅ Verification

After installation, you can:
- **Find it in Spotlight**: Press `Cmd+Space` and type "Launch EE GUI"
- **Add to Dock**: Drag from Applications folder to Dock
- **Launch from Launchpad**: Look for "Export for My Efficient Elements"

## 🔧 How It Works

The app uses **absolute paths** to your project directory, so it will work from anywhere:
- **Project Location**: `/Users/upmanyuskand/repos/ee_my_elements`
- **Virtual Environment**: Automatically activated
- **Dependencies**: Automatically checked and installed if needed

## 🛠️ Troubleshooting

### If the app doesn't launch:
1. **Check project location**: Ensure the project is still at `/Users/upmanyuskand/repos/ee_my_elements`
2. **Check permissions**: The app might need permission to run
   ```bash
   # If needed, reset permissions
   chmod +x "/Applications/Launch EE GUI.app/Contents/MacOS/launch_ee_gui"
   ```
3. **Check virtual environment**: Ensure `venv` folder exists in the project

### If you moved the project:
If you move the project to a different location, you'll need to update the app:
1. Edit the app's executable: `/Applications/Launch EE GUI.app/Contents/MacOS/launch_ee_gui`
2. Update the `PROJECT_DIR` variable to the new location
3. Or recreate the app from the new project location

## 🎯 Alternative: Create Symlink

If you prefer to keep the app in the project folder but access it from Applications:
```bash
# Create a symbolic link in Applications
ln -s "/Users/upmanyuskand/repos/ee_my_elements/Launch EE GUI.app" "/Applications/Launch EE GUI.app"
```

This way, the app stays in your project folder but appears in Applications.

## 🔄 Updates

When you update the project or the launch scripts:
- **If using copy method**: Copy the updated app to Applications again
- **If using symlink method**: No action needed, it will use the updated version automatically

---

**Enjoy your portable Export for My Efficient Elements launcher! 🎉**
