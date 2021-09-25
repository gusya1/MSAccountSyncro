import datetime
from settings import *
from tkinter import *

from MSApi.MSApi import MSApi, error_handler
from MSApi.exceptions import MSApiException
from MSApi.properties import Expand, Filter
from threading import Thread
import tkinter.messagebox as mb
from tkcalendar import Calendar


class MainWindow:

    def __init__(self):
        self.ui = Tk()
        self.ui.geometry('500x500')
        self.ui.grid_columnconfigure(0, weight=1)
        self.ui.resizable(True, True)
        self.ui.title("AccountSyncro")

        self.generate_btn = Button(self.ui, text="Sync", command=self.on_sync_button)
        self.generate_btn.grid(column=0, row=1, columnspan=2, sticky=NSEW)

        self.text_browser = Text(self.ui)
        self.text_browser.config(state=DISABLED)
        self.text_browser.grid(column=0, row=2, sticky=NSEW)
        scroll = Scrollbar(command=self.text_browser.yview)
        scroll.grid(column=1, row=2, sticky=NSEW)
        self.text_browser.config(yscrollcommand=scroll.set)

        self.calendar = Calendar(self.ui)
        self.calendar.grid(column=0, row=0, columnspan=2, sticky=NSEW)
        self.ui.mainloop()

    def print_text(self, string):
        self.text_browser.configure(state='normal')
        self.text_browser.insert(END, f"{string}\n")
        self.text_browser.configure(state='disabled')

    def clear_text(self):
        self.text_browser.configure(state='normal')
        self.text_browser.delete(0.0, END)
        self.text_browser.configure(state='disabled')

    def on_sync_button(self):
        th = Thread(target=self.synchronize)
        th.start()

    def synchronize(self):
        try:
            self.clear_text()
            self.generate_btn.configure(state='disabled')

            date_obj = datetime.datetime.strptime(self.calendar.get_date(), "%m/%d/%y")
            date_filter = Filter.gte('deliveryPlannedMoment', date_obj.strftime('%Y-%m-%d'))
            date_filter += Filter().lt('deliveryPlannedMoment',
                                       (date_obj + datetime.timedelta(days=1)).strftime('%Y-%m-%d'))

            self.print_text(f"Reading accounts...")
            accounts_meta_dict = {}
            for organization in MSApi.gen_organizations():
                if organization.get_name() == organization_name:
                    for account in organization.gen_accounts():
                        accounts_meta_dict[account.get_account_number()] = account.get_meta()
                    break
            else:
                raise RuntimeError("Organization not found!")

            if len(accounts_meta_dict) == 0:
                raise RuntimeError("Accounts list is empty!")

            updatable_customerorder_list = []

            for state_name, account_name in state_dict.items():
                self.print_text(f"Status \"{state_name}\" checking...")
                account_meta = accounts_meta_dict.get(account_name)
                if account_meta is None:
                    raise RuntimeError(f"Account name \"{account_name}\" not found")
                filters = Filter.eq("state.name", state_name) + date_filter
                for customer_order in MSApi.gen_customer_orders(expand=Expand("state"), filters=filters):
                    org_acc = customer_order.get_organization_account()
                    if org_acc is not None:
                        if org_acc.get_meta() == account_meta:
                            continue

                    self.print_text(f"\tCustomer order \"{customer_order.get_name()}\" will be changed")
                    updatable_customerorder_list.append(
                        {
                            'meta': customer_order.get_meta().get_json(),
                            'organizationAccount': {'meta': account_meta.get_json()}
                        })

            if len(updatable_customerorder_list) == 0:
                raise RuntimeError("Customer orders not found!")

            self.print_text(f"Send to MoySklad...")
            response = MSApi.auch_post("entity/customerorder", json=updatable_customerorder_list)
            error_handler(response)
            self.print_text("Success!")
            mb.showinfo("Info", "Success!")
        except RuntimeError as exception:
            self.print_text(str(exception))
            mb.showerror("Error", str(exception))
        except MSApiException as exception:
            self.print_text(str(exception))
            mb.showerror("Error", str(exception))
        self.generate_btn.configure(state='normal')


if __name__ == "__main__":
    try:
        MSApi.login(login, password)
        window = MainWindow()
    except RuntimeError as exception:
        mb.showerror("Error", str(exception))
    except MSApiException as exception:
        mb.showerror("Error", str(exception))



