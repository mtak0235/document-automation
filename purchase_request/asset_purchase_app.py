import tkinter as tk

class Clipboard:
    def __init__(self):
        pass
    
    def copy_text(self, text):
        pass
    
class Excel:
    def __init__(self):
        pass
    
    def save(self, data):
        pass

from tkinter import ttk

class View:
    def __init__(self, master):
        self.master = master
        self.master.title("Asset Purchase App")
         # define font, style
        self.set_font()
        
        #generate labels and fields
            
        # generate treeview
        self.input_headers = ["목적", "내용", "구입업체", "품명", "수량", "단가"]
        self.tree_headers = ["ID"] + self.input_headers + ["구분", "금액"]
        self .tree =  ttk.Treeview(self.master, columns=self.tree_headers, show="headings")
        for i in range(len(self.tree_headers)):
            self.tree.heading(f"#{i+1}", text=self.tree_headers[i])
        
        # Treeview column width and alignment
        self.tree.column("ID", width=50)
        self.tree.column("목적", width=150)
        self.tree.column("내용", width=150)
        self.tree.column("구입업체", width=150)
        self.tree.column("구분", width=70)
        self.tree.column("품명", width=150)
        self.tree.column("수량", width=70) 
        self.tree.column("단가", width=70)
        self.tree.column("금액", width=70)
        
         # Add selection binding
        self.tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        
        #data access
        self.datas = Assets()
        self.asset_manager = AssetManager(self.datas)
        
        # services
        self.clipboard = Clipboard()
        self.excel = Excel()
    
    def render_view(self):
        
        # render labels and fields
        tk.Label(self.master, text="").grid(row=0, pady=10)
        
        tk.Label(self.master, text="목적\t:").grid(row=1, column=0, sticky="e", pady=5)
        self.purpose_entry = tk.Entry(self.master, width=50)
        self.purpose_entry.grid(row=1, column=1, sticky="w")
        
        tk.Label(self.master, text="내용\t:").grid(row=2, column=0, sticky="e", pady=5)
        self.inner_content_entry = tk.Text(self.master, width=50, height=5)
        self.inner_content_entry.grid(row=2, column=1, sticky="w")
        
        tk.Label(self.master, text="구입업체\t:").grid(row=3, column=0, sticky="e", pady=5)
        self.selling_company_entry = tk.Entry(self.master, width=50)
        self.selling_company_entry.grid(row=3, column=1, sticky="w")
        
        tk.Label(self.master, text="품명\t:").grid(row=4, column=0, sticky="e", pady=5)
        self.item_name_entry = tk.Entry(self.master, width=50)
        self.item_name_entry.grid(row=4, column=1, sticky="w")
        
        tk.Label(self.master, text="수량\t:").grid(row=5, column=0, sticky="e", pady=5)
        self.quantity_entry = tk.Entry(self.master, width=50)
        self.quantity_entry.grid(row=5, column=1, sticky="w")
        
        tk.Label(self.master, text="단가\t:").grid(row=6, column=0, sticky="e", pady=5)
        self.unit_price_entry = tk.Entry(self.master, width=50)
        self.unit_price_entry.grid(row=6, column=1, sticky="w")
        
        #render buttons
        
        manipulate_buttons_frame = tk.Frame(self.master)
        manipulate_buttons_frame.grid(row=7, column=0, columnspan=2)
        
        
        tk.Button(manipulate_buttons_frame, text="행 추가", command=lambda: self.add_row()).pack(side="left", padx=5)
        tk.Button(manipulate_buttons_frame, text="행 수정", command=lambda: self.modify_row()).pack(side="left", padx=5)
        tk.Button(manipulate_buttons_frame, text="행 삭제", command=lambda: self.delete_row()).pack(side="left", padx=5)
        
        
        result_buttons_frame = tk.Frame(self.master)
        result_buttons_frame.grid(row=9, column=0, columnspan=2, pady=10)
        
        #TODO : deal data
        clipboard_datas = None
        tk.Button(result_buttons_frame, text="텍스트 클립보드에 복사", command=lambda: self.copy_text_to_clipboard(clipboard_datas)).pack(side="left", padx=5)
        
        # TODO : deal data
        excel_datas = None
        tk.Button(result_buttons_frame, text="차트 엑셀로 열기", command=lambda: self.save_excel(excel_datas)).pack(side="left", padx=5)
        
        # render treeview
        self.tree.grid(row=8, column=0, columnspan=2, pady=5)
    
    def clear_inputs(self):
        """Clear all input fields"""
        self.purpose_entry.delete(0, tk.END)
        self.inner_content_entry.delete('1.0', tk.END)
        self.selling_company_entry.delete(0, tk.END)
        self.item_name_entry.delete(0, tk.END)
        self.quantity_entry.delete(0, tk.END)
        self.unit_price_entry.delete(0, tk.END)

    def on_tree_select(self, event):
        """Handle treeview selection event"""
        selected_items = self.tree.selection()
        if not selected_items:
            return
            
        # Get the selected item's values
        selected_values = self.tree.item(selected_items[0])['values']
        
        # Clear current inputs
        self.clear_inputs()
        
        # Fill the input fields with selected values
        self.purpose_entry.insert(0, selected_values[0])
        self.inner_content_entry.insert('1.0', selected_values[1])
        self.selling_company_entry.insert(0, selected_values[2])
        self.item_name_entry.insert(0, selected_values[3])
        self.quantity_entry.insert(0, selected_values[4])
        self.unit_price_entry.insert(0, selected_values[5])
    
    def add_row(self):
        data = Asset.builder() \
        .purpose(self.purpose_entry.get()) \
        .content(self.inner_content_entry.get("1.0", tk.END)) \
        .vendor(self.selling_company_entry.get()) \
        .item_name(self.item_name_entry.get()) \
        .quantity(int(self.quantity_entry.get())) \
        .unit_price(int(self.unit_price_entry.get()) )\
        .build()
        # save data
        self.asset_manager.add_asset(data)
        # add row to treeview
        tree_data = [data.id, data.purpose, data.content, data.vendor, data.item_name, data.quantity, data.unit_price, data.category, data.total_price]
        self.tree.insert("", "end", values=tree_data)
        # Clear input fields after adding
        self.clear_inputs()

        
    def modify_row(self):
        selected_item = self.tree.selection()
        if not selected_item:
            tk.messagebox.showwarning("Warning", "Please select a record to modify")
            return

        # Get values from the selected item
        selected_values = self.tree.item(selected_item[0])["values"]
        asset_id = selected_values[0]
        
        # Create the new modified asset
        modified_asset = Asset.builder() \
            .id(asset_id) \
            .purpose(self.purpose_entry.get()) \
            .content(self.inner_content_entry.get("1.0", tk.END)) \
            .vendor(self.selling_company_entry.get()) \
            .item_name(self.item_name_entry.get()) \
            .quantity(int(self.quantity_entry.get())) \
            .unit_price(int(self.unit_price_entry.get())) \
            .build()
        
        # Modify the asset in the AssetManager
        self.asset_manager.modify_asset(asset_id, modified_asset)
        
        # Update the treeview with modified data
        tree_data = [modified_asset.id, modified_asset.purpose, modified_asset.content, 
                    modified_asset.vendor, modified_asset.item_name, 
                    modified_asset.quantity, modified_asset.unit_price, 
                    modified_asset.category, modified_asset.total_price]
        self.tree.item(selected_item, values=tree_data)
        
        # Clear input fields after modifying
        self.clear_inputs()
    
    def delete_row(self):
       selected_items = self.tree.selection()
       if not selected_items:
            tk.messagebox.showwarning("Warning", "Please select a record to delete")
            return
            
       if tk.messagebox.askyesno("Delete", "Are you sure you want to delete this record?"):
            for selected_record in selected_items:
                # Get the ID of the selected item
                asset_id = self.tree.item(selected_record)['values'][0]
                
                # Remove from asset manager using ID
                self.asset_manager.remove_asset(asset_id)
                
                # Remove from treeview
                self.tree.delete(selected_record)
            
            self.clear_inputs()
            
    
    def get_buttons(self):
        return self.buttons
    
    def set_font(self):
        pass
    
    def copy_text_to_clipboard(self, data):
        self.clipboard.copy_text(data)
        
    def save_excel(self, data):
        self.excel.save(data)
    
class Asset:
    def __init__(self):
        self.id = None
        self.purpose = None
        self.content = None
        self.vendor = None
        self.item_name = None
        self.quantity = 1
        self.unit_price = None
        self.category = "신규 구매"
        self.total_price = None
    
    def __eq__(self, other):
        if not isinstance(other, Asset):
            return False
        return self.id == other.id
    
    def __hash__(self):
        return hash(self.id)
    
    @staticmethod
    def builder():
        return Asset.Builder()

    class Builder:
        def __init__(self):
            self.asset = Asset()

        def purpose(self, purpose):
            self.asset.purpose = purpose
            return self

        def content(self, content):
            self.asset.content = content
            return self

        def vendor(self, vendor):
            self.asset.vendor = vendor
            return self

        def item_name(self, item_name):
            self.asset.item_name = item_name
            return self

        def quantity(self, quantity):
            self.asset.quantity = int(quantity)
            return self

        def unit_price(self, unit_price):
            self.asset.unit_price = int(unit_price)
            self.asset.total_price = int(self.asset.quantity) * int(self.asset.unit_price)
            return self

        def category(self, category):
            self.asset.category = category
            return self
        
        def id(self, id):
            self.asset.id = id
            return self

        def build(self):
            return self.asset

class Assets:
    def __init__(self):
        self.assets = []
        self._next_id = 1
    
    def generate_id(self):
        id = self._next_id
        self._next_id += 1
        return id
    
    def add_asset(self, asset):
        if isinstance(asset, Asset):
            if asset.id is None:
                asset.id = self.generate_id()
            self.assets.append(asset)
        else:
            raise TypeError("Only Asset objects can be added")

    def remove_asset(self, asset_id):
        self.assets = [asset for asset in self.assets if asset.id != asset_id]
        
    def get_asset_by_id(self, asset_id):
        for asset in self.assets:
            if asset.id == asset_id:
                return asset
        return None
    
    def get_all_assets(self):
        return self.assets

    def get_asset_by_index(self, index):
        return self.assets[index]

    def get_total_value(self):
        return sum(asset.total_price for asset in self.assets if asset.total_price is not None)

    def get_assets_by_category(self, category):
        return [asset for asset in self.assets if asset.category == category]

    def get_assets_by_vendor(self, vendor):
        return [asset for asset in self.assets if asset.vendor == vendor]

    def __len__(self):
        return len(self.assets)

    def __iter__(self):
        return iter(self.assets)

class AssetManager:
    def __init__(self, assets):
        self.assets = assets
    
    def add_asset(self, asset):
        self.assets.add_asset(asset)
    
    def remove_asset(self, asset_id):
        self.assets.remove_asset(asset_id)
    
    def modify_asset(self, asset_id, new_asset):
        # Remove old asset and add the modified one
        self.remove_asset(asset_id)
        self.add_asset(new_asset)
    
class AssetPurchaseApp:
    def __init__(self, master):
        
        self.master = master
        
        self.view = View(self.master)
        self.view.render_view()