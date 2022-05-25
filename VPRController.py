from enum import Enum
from re import T
from pywinauto import application, ElementAmbiguousError
from time import sleep

class VPRTypes(Enum):
    VPR_FORM = 0
    VPR_X12 = 1


class VPRController:
    def __init__(self) -> None:
        self.applicationReference = application.Application()
        try:
            self.reference = self.applicationReference.connect(title_re="Viewer Printer")
            self.vpReference = self.reference.ViewerPrinter
        except:
            raise Exception("Could not connect to Viewer Printer...\nMaybe not open?\nError Dialouge open?")

        self.currentTabState = self._determinePage()

    def _determinePage(self) -> None:
        edit3Text = self.vpReference.Edit3.window_text()
        if len(edit3Text) == 12 or len(edit3Text) == 0:
            print("VPR Form Selected")
            print("If wrong - Call invertTabState().")
            return VPRTypes.VPR_FORM
        else:
            print("VPR X12 Selected")
            print("If wrong - Call invertTabState().")
            return VPRTypes.VPR_X12

    def logCurrentState(self) -> None: 
        print(f"Current Tab State: {self.currentTabState}")

    def invertTabState(self) -> None:
        '''
        Cycles tab state
        '''
        if self.currentTabState == VPRTypes.VPR_FORM:
            self.currentTabState = VPRTypes.VPR_X12
            print("Tab State: VPR X12")
        else:
            self.currentTabState = VPRTypes.VPR_FORM
            print("Tab State: VPR Form")

    def updateClaimNum(self, claimNum:str, search:bool=False, checkWrongClaim:bool=True, timeout:float|int=3, log:bool=False) -> bool:
        '''
        Updates field labeled "DCN" with arg -> claimNum
        args:
            claimNum:str -> updated claim number
            search:bool=False -> click on search button
            log:bool=False -> print log information
        return:
            True: claim number updated and no error
            False: error updating claim number/claim number not updated
        '''        
        for _ in range(20): # Delete all in field before updating
            self.vpReference.Edit0.type_keys("{BACKSPACE}")

        self.vpReference.Edit0.type_keys(claimNum)
        if search: self.vpReference.SearchButton.click()
        if log: print("Claim Number Updated")
        if checkWrongClaim:
            sleep(timeout)
            try: # Attempt to connect to see if claim number failed to search
                self.applicationReference.connect(title_re="Viewer Printer") 
            except ElementAmbiguousError:
                self.reference.ViewerPrinter.OK.click() # Click OK on error Dialouge
                return False
        
        return True
    
    def returnX12(self, printOut:bool=False, log:bool=False) -> str:
        '''
        returns x12 data
        args:
            printOut:bool=False -> print x12 data to console
            log:bool=False -> print log information
        return:
            x12:str -> x12 data
        '''
        self.openX12()

        x12 = self.vpReference.Edit3.window_text()
        if printOut: print(x12)
        if log: print("X12 Data Procured")
        return x12

    def openX12(self) -> None:
        '''
        opens the x12 data tab
        '''
        if self.currentTabState == VPRTypes.VPR_X12:
            return

        self.vpReference.Tab1TabControl.click()
        self.vpReference.type_keys("{VK_LEFT}")
        self.currentTabState = VPRTypes.VPR_X12

    def openForm(self) -> None:
        '''
        opens the form data tab
        '''
        if self.currentTabState == VPRTypes.VPR_FORM:
            return

        self.vpReference.Tab1TabControl.click()
        self.vpReference.type_keys("{VK_RIGHT}")
        self.currentTabState = VPRTypes.VPR_FORM

if __name__ == "__main__":
    '''
    Testing Information
    '''
    controller = VPRController()
    getClaim = controller.updateClaimNum("305631381800", True, True)
    if getClaim:
        x12 = controller.returnX12()
        controller.openX12()
        controller.openForm()
        print(x12 != True) # Check if x12 string is empty
    else:
        print("Claim Failed to Load")