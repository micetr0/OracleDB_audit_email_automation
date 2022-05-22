[void][System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")
#add-type -AssemblyName System.Data.OracleClient

function Run-OracleProcedure
{
 <#*******************************************************************************
 Purpose: invoke oracle stored procedure

 Return: null

 Dependency: none
    
 Modifications
 Date           Author          Description                     
 ---------------------------------------------------------
 11-Mar-2021    William Hu     Initial version
 *******************************************************************************#> 

    [CmdletBinding()]
    Param
    (
        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OracleUser,

        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OraclePassword,
        
        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OracleHost,

        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OracleServericeName,

        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OracleProcedure
    )
              
    Process
    {    
        $connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$OracleHost)(PORT=1521)) (CONNECT_DATA=(SERVICE_NAME=$OracleServericeName)));User Id=$OracleUser;Password=$OraclePassword;"
        try 
        {
            $con = new-object system.data.oracleclient.oracleconnection($connectionString)    
            $cmd = new-object System.Data.OracleClient.OracleCommand;
            $cmd.Connection = $con
            $cmd.CommandText = $OracleProcedure 
            $cmd.CommandType = [System.Data.CommandType]::StoredProcedure;
            $con.open()
            $cmd.ExecuteNonQuery() |Out-Null 
        }
        catch 
        {
            Write-Error (“Database Exception: {0}`n{1}” -f `
                $con.ConnectionString, $_.Exception.ToString())                    
        }
        finally
        {
            if ($con.State -eq ‘Open’) { $con.close() }            
        }
    }

}

function Run-OracleSQLQuery
{
 <#*******************************************************************************
 Purpose: return result from oracle sql query

 Return: data.datatable 

 Dependency: none
    
 Modifications
 Date           Author          Description                     
 ---------------------------------------------------------
 11-Mar-2021    William Hu     Initial version
 *******************************************************************************#> 

    [CmdletBinding()]
    Param
    (
        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OracleUser,

        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OraclePassword,
        
        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OracleHost,

        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OracleServericeName,

        [string]
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        $OracleSQLQuery
    )

    Process
    {    
        $connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$OracleHost)(PORT=1521)) (CONNECT_DATA=(SERVICE_NAME=$OracleServericeName)));User Id=$OracleUser;Password=$OraclePassword;"
        try 
        {
            $con = new-object system.data.oracleclient.oracleconnection($connectionString)    
            $cmd = new-object System.Data.OracleClient.OracleCommand;
            $con.open()
            $list_set = new-object system.data.dataset
            $list_adapter = new-object system.data.oracleclient.oracledataadapter($OracleSQLQuery, $con)
            $list_adapter.Fill($list_set) | Out-Null
            $list_table = new-object system.data.datatable
            $list_table = $list_set.Tables[0]

            return $list_table
        }
        catch 
        {
            Write-Error (“Database Exception: {0}`n{1}” -f `
                $con.ConnectionString, $_.Exception.ToString())    
        }
        finally
        {
            if ($con.State -eq ‘Open’) { $con.close() }            
        }
    }

}