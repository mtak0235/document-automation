# main.py
import tkinter as tk
from asset_purchase_app import AssetPurchaseApp
from ui import create_ui

if __name__ == "__main__":
    root = tk.Tk()
    assetPurchaseApp = AssetPurchaseApp(root)
    root.mainloop()