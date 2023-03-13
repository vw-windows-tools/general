class BaseClass{

    [int] $Id
    hidden static [int16] $InstancesCount

    Watcher() {
        $type = $this.GetType()

        if ($type -eq [BaseClass])
        {
            throw("Class $type must be inherited")
        }
    }

    [string]ToString(){
        throw("Must Override Method")
    }

}

class InheritedClass : BaseClass {

    #region Properties
    [int]$SerialNumber
    [string]$MyString
    hidden [array]$SomeList = @("Item 1","Item 2","Item 3")
    hidden [string]$Mode
    #endregion

    #region Constructors
    InheritedClass() {
        Throw "SerialNumber is required"
    }

    InheritedClass([UInt32]$SerialNumber) {
        $This.InitFromSerialNumber($SerialNumber)
    }

    InheritedClass([UInt32]$SerialNumber, [string]$Mode) {
        $This.InitFromSerialNumber($SerialNumber)
        $This.SetMode($Mode)
    }
    #endregion

    #region Setters
    [void]SetMode([string]$Mode){
        Switch ($Mode) {
            "easy" { 
                $This.Mode = "easy"
                Write-Verbose "Mode set to ""Easy"""
            }
            "hard" { 
                $This.Mode = "hard"
                Write-Verbose "Mode set to ""Hard""" 
            }
            Default {Throw "Mode unknown"}
        }
    }
    #endregion

    #region Getters
    [string]GetMode(){
        Switch ($This.Mode) {
            "easy" { return "Easy" }
            "hard" { return "Hard" }
        }
        return "Runtime not set"
    }
    #endregion

    #region Other methods
    hidden [void]InitFromSerialNumber($CommandLine) {
        [InheritedClass]::InstancesCount++ # Increment count of watchers
        $This.Id = [InheritedClass]::InstancesCount # Current number of watchers is assigned to this as an ID
        Write-Verbose "Instance $($this.Id) initialized"
    }

    [string]ToString(){
        return "Id = $($this.Id) | ProcessId = $($this.ProcessId) | Version = $($this.Version)"
    }
    #endregion

}

$VerbosePreference="continue"

# New Object 1
$TestObject1 = [InheritedClass]::new(8910,'easy')

# New Object 2
$serial = 34567
$mode = "hard"
$TestObject2 = New-Object -TypeName InheritedClass -ArgumentList $serial
$TestObject2.GetMode()
$TestObject2.SetMode($mode)