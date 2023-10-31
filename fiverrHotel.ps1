$allGuests = Import-Excel C:\Users\herngyih\Documents\Fiverr\Hotel\fiverr.xlsx -WorksheetName Attendees    
 
$rooms = Import-Excel C:\Users\herngyih\Documents\Fiverr\Hotel\fiverr.xlsx -WorksheetName Rooms

<#
    ALLGUESTS FUNCTIONS

        1.1 Put them in a 2d array (line 33) - Completed
    
        1.2 get total roomless members (line 46) - Completed

        1.3 set hasRoom (line 56) - Completed
            - set guest.hasRoom to True once assigned to a room

       

    ROOM FUNCTIONS
        
        2.1 get vacant rooms (line 75) - Completed
            - from the rooms,pick the vacant one
    
        2.2 get perfect room (line 87) - Completed
            - from the vacant rooms, pick the one which has cap eq roomless members
        
        2.3 get smaller room (line 113) - Completed
            - for the vacant rooms, pick the one which cap < than roomless members 

        2.4 set vacant (line 138) - Completed
            - Set vacant to false once it is filled
    
    TEST AREA

        1 (line 151)

        2 (line 175)

        3 (line 199)
#>

#Start of 1.1
    $allGuests2d = New-Object System.Collections.ArrayList

    $guestsUniqueGroup = $allGuests.Group| select -Unique

    for($i=0; $i -lt $guestsUniqueGroup.Count; $i++){
        
        $guestsByGroup = $allGuests|Where-Object{$_.Group -eq $guestsUniqueGroup[$i]}
    
        $allGuests2d.Add(@($guestsByGroup))|Out-Null
    }    
#End of 1.1

#Start of 1.2
    
    #return total number of roomless people in the group
    function getRoomlessMembers($groupIndex,$allGuests2d){

        return ($allGuests2d[$groupIndex]|Where-Object{$_.hasRoom -eq $false}).Count
    }

#End of 1.2    

#Start of 1.3
    
    #set a group member hasRoom to $true and assign the room number
    function setHasRoom($row,$allGuests2d,$roomNumber){

        $roomlessMembers = $allGuests2d[$row]|Where-Object{$_.hasRoom -eq $false}
        $roomlessMembers[0].hasRoom = $true
        $roomlessMembers[0].roomNumber = $roomNumber
    }
#End of 1.3



#Start of 2.1

    #return an array of vacant rooms
    function getVacantRooms($rooms){

        $vacantRooms = $rooms|Where-Object{$_.Vacant -eq $true} 

        return $vacantRooms
    }

#End of 2.1   

#Start of 2.2

    #return a vacant room object which capacity is equal to the number of roomless members
    function getPerfectRoom($vacantRooms,$totalRoomlessMembers){

        $perfectRoomObj = $null

        foreach($vacantRoom in $vacantRooms){

            if($vacantRoom.capacity -eq $totalRoomlessMembers){
                
                $perfectRoomObj= $vacantRoom
                
                break
            }

            
        }

            return $perfectRoomObj
        
    }

#End of 2.2


#start of 2.3 
     
    #return a vacant room number which capacity is less than the number of roomless members
    function getSmallRoom($vacantRooms,$totalRoomlessMembers){

        $smallRoomNumber = $null

        foreach($vacantRoom in $vacantRooms){

            if($vacantRoom.capacity -lt $totalRoomlessMembers){
                
                $smallRoomNumber = $vacantRoom
                
                break
            }

            
        }

            return $smallRoomNumber
        
     }
            
#

#2.4 set vacant

    #set a room vacant property to false
    function setVacant($rooms,$roomNumber){
        
        ($rooms|Where-Object{$_.Number -eq $roomNumber}).vacant = $false
        
    }
            
#



#Start of test area 1

    <#    
    #Perfect room function tester
    
    for($k=0; $k -lt $allGuests2d.Count; $k++){

        $perfectRoomObj = getPerfectRoom (getVacantRooms $rooms) (getRoomlessMembers $k $allGuests2d)

        if($perfectRoomObj -ne $null){

                Write-Host -NoNewline "Room: Perfect "  
                Write-Host -NoNewline "Group:" ($allGuests2d[$k].Group|select -Unique)
                Write-Host -NoNewline " Roomless:"(getRoomlessMembers $k $allGuests2d)
                Write-Host -NoNewline " Number:" $perfectRoomObj.number
                Write-Host " Capacity:" $perfectRoomObj.capacity
         }
    }
    #>
        
          
#End of test area 1


#Start of test area 2
    
   <#    
    #Small room function tester

    for($k=0; $k -lt $allGuests2d.Count; $k++){

        $smallRoomObj = getSmallRoom (getVacantRooms $rooms) (getRoomlessMembers $k $allGuests2d)

        if($smallRoomObj -ne $null){

                Write-Host -NoNewline "Room: Small "  
                Write-Host -NoNewline "Group:" ($allGuests2d[$k].Group|select -Unique)
                Write-Host -NoNewline " Roomless:"(getRoomlessMembers $k $allGuests2d)
                Write-Host -NoNewline " Number:" $smallRoomObj.number
                Write-Host " Capacity:" $smallRoomObj.capacity
        }
    }
   
   #>    

#End of test area 2



#Start of test area 3

   
    for($k=0; $k -lt $allGuests2d.Count; $k++){

        $continue = $true

        do{
            $perfectRoomObj = getPerfectRoom (getVacantRooms $rooms) (getRoomlessMembers $k $allGuests2d)            
            
            
             if($perfectRoomObj -ne $null){

                for($l=0 ; $l -lt $perfectRoomObj.capacity; $l++){
                  
                    setHasRoom $k $allGuests2d $perfectRoomObj.Number
                  
                 }

                $perfectRoomObj.Vacant = $false
            }       
            
            else{
                
                $smallRoomObj = getSmallRoom (getVacantRooms $rooms) (getRoomlessMembers $k $allGuests2d)

                if($smallRoomObj -ne $null){

                    for($l=0 ; $l -lt $smallRoomObj.capacity; $l++){
                        
                       setHasRoom $k $allGuests2d $smallRoomObj.Number
                    }

                    $smallRoomObj.Vacant = $false

                 }
            
                else{
                
                    $continue = $false
                }
            } 
        }

        while($continue)
    }
 

#End of test area 3


#Start of test area 4

    $allGuests2d|Export-Excel fiverrHotelResult.xlsx
#End of test area 4