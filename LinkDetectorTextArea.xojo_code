#tag Class
Protected Class LinkDetectorTextArea
Inherits TextArea
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  
		  if linkedWords <> nil and boundedWord <> nil then
		    
		    
		    if linkedWords.HasKey(mboundedWord.Left) then
		      
		      dim p as new pair(mboundedWord.Left,linkedWords.Value(mboundedWord.Left))
		      
		      for Each item as LinkDetectorTextArea.URLTextActionDelegate in observers
		        
		        
		        item.Invoke p
		        
		        
		      next
		      
		      
		    end if
		    
		    Return true
		    
		  end if
		  
		  
		  Return RaiseEvent mousedown(x,y)
		End Function
	#tag EndEvent

	#tag Event
		Sub MouseEnter()
		  
		  lenght = me.Text.Len
		  
		  RaiseEvent mouseEnter
		End Sub
	#tag EndEvent

	#tag Event
		Sub MouseExit()
		  
		  RaiseEvent MouseExit
		End Sub
	#tag EndEvent

	#tag Event
		Sub MouseMove(X As Integer, Y As Integer)
		  
		  if previousBoundedWord <> nil then
		    
		    setStyleForWord(previousBoundedWord, false)
		    previousBoundedWord = nil
		    
		  end if
		  
		  
		  dim boundLeft, boundRight as integer
		  
		  dim startPosition as integer = me.CharPosAtXY(x,y)
		  
		  if CharPosAtxy(x,y) >= me.Text.len then
		    
		    
		    if boundedWord <> nil then setStyleForWord(boundedWord,false)
		    mboundedWord = nil
		    
		  elseif me.text.mid(startPosition,1) <> chr(32) and me.Text.mid(startPosition,1) <> EndOfLine then
		    
		    for n as integer = startPosition DownTo 1
		      
		      if me.Text.mid(n-1,1) = chr(32) or me.Text.mid(n-1,1) = EndOfLine then
		        boundLeft = n
		        exit
		      end
		      
		    next
		    
		    for n as integer = startPosition to lenght
		      
		      if me.Text.mid(n+1,1) = chr(32) or me.Text.mid(n+1,1) = EndOfLine or n = lenght then
		        boundRight = n+1
		        exit
		      end
		      
		    next
		    
		  end
		  
		  dim isolatedWord as string = me.Text.mid(boundLeft, boundRight - boundLeft)
		  
		  dim check as pair = wordInDictionary( isolatedWord, boundleft, boundRight )
		  
		  if check  <> nil then 
		    
		    mboundedWord  = check
		    
		    if previousBoundedWord = nil then previousBoundedWord = mboundedWord
		    
		    setStyleForWord(previousBoundedWord, true)
		    
		    
		  end if
		  
		  
		  RaiseEvent mouseMove(X, Y)
		End Sub
	#tag EndEvent

	#tag Event
		Sub Open()
		  me.MouseCursor = System.cursors.standardpointer
		  
		  RaiseEvent open
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub deleteObserver(observer as LinkDetectorTextArea.URLTextActionDelegate)
		  if observer <> nil then
		    
		    dim n as integer = observers.IndexOf(observer)
		    
		    observers.Remove(n)
		    
		  end
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub registerObserver(observer as URLTextActionDelegate)
		  if observer <> nil then observers.Append observer
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub setDictionary(d as Dictionary)
		  linkedWords = d
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub setStyleForWord(word as pair, mode as Boolean)
		  dim cStart, cEnd, sStart as integer
		  
		  sStart = me.SelStart
		  
		  cStart = NthField(word.Right.TextValue,"-",1).Val
		  cEnd = NthField(word.Right.TextValue,"-",2).val - cStart
		  
		  me.SelStart = cStart - 1
		  me.SelLength = cEnd
		  
		  me.SelUnderline = mode
		  
		  me.SelStart = cStart
		  me.SelLength = 0
		End Sub
	#tag EndMethod

	#tag DelegateDeclaration, Flags = &h0
		Delegate Sub URLTextActionDelegate(boundedWord as pair)
	#tag EndDelegateDeclaration

	#tag Method, Flags = &h21
		Private Function wordInDictionary(word as string, leftposition as integer, RightPosition as integer) As pair
		  if linkedWords.HasKey(word) then Return new pair(word,leftposition.ToText+"-"+RightPosition.ToText)
		  
		  dim blu as integer
		  
		  dim dictionaryKeys() as variant = linkedWords.Keys
		  
		  for each thekey as string in dictionaryKeys
		    
		    if thekey.InStr(word) > 0 then
		      
		      blu = me.text.instr(thekey)
		      
		      if leftposition >= blu and leftposition <= (blu + thekey.Len) then
		        
		        dim finalposition as integer =blu+thekey.len
		        Return new pair(thekey,blu.ToText+"-"+finalposition.totext)
		        
		      end if
		      
		    end
		    
		  next
		  
		  Return nil
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event MouseDown(x as integer, y as integer) As Boolean
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MouseEnter()
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MouseExit()
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MouseMove(X as integer, Y as integer)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event Open()
	#tag EndHook


	#tag ComputedProperty, Flags = &h0
		#tag Getter
			Get
			  return mboundedWord
			End Get
		#tag EndGetter
		boundedWord As pair
	#tag EndComputedProperty

	#tag Property, Flags = &h21
		Private lenght As Integer
	#tag EndProperty

	#tag Property, Flags = &h21
		Private linkedWords As Dictionary
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mboundedWord As pair
	#tag EndProperty

	#tag Property, Flags = &h21
		Private observers() As URLTextActionDelegate
	#tag EndProperty

	#tag Property, Flags = &h21
		Private previousBoundedWord As Pair
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="AcceptTabs"
			Visible=true
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Alignment"
			Visible=true
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
			EditorType="Enum"
			#tag EnumValues
				"0 - Default"
				"1 - Left"
				"2 - Center"
				"3 - Right"
			#tag EndEnumValues
		#tag EndViewProperty
		#tag ViewProperty
			Name="AutoDeactivate"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="AutomaticallyCheckSpelling"
			Visible=true
			Group="Behavior"
			InitialValue="True"
			Type="boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="BackColor"
			Visible=true
			Group="Appearance"
			InitialValue="&hFFFFFF"
			Type="Color"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Bold"
			Visible=true
			Group="Font"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Border"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DataField"
			Visible=true
			Group="Database Binding"
			Type="String"
			EditorType="DataField"
		#tag EndViewProperty
		#tag ViewProperty
			Name="DataSource"
			Visible=true
			Group="Database Binding"
			Type="String"
			EditorType="DataSource"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Enabled"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Format"
			Visible=true
			Group="Appearance"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Height"
			Visible=true
			Group="Position"
			InitialValue="100"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HelpTag"
			Visible=true
			Group="Appearance"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="HideSelection"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			Type="Integer"
			EditorType="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Italic"
			Visible=true
			Group="Font"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LimitText"
			Visible=true
			Group="Behavior"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LineHeight"
			Visible=true
			Group="Appearance"
			InitialValue="0"
			Type="Double"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LineSpacing"
			Visible=true
			Group="Appearance"
			InitialValue="1"
			Type="Double"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockBottom"
			Visible=true
			Group="Position"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockLeft"
			Visible=true
			Group="Position"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockRight"
			Visible=true
			Group="Position"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="LockTop"
			Visible=true
			Group="Position"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Mask"
			Visible=true
			Group="Behavior"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Multiline"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ReadOnly"
			Visible=true
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ScrollbarHorizontal"
			Visible=true
			Group="Appearance"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ScrollbarVertical"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Styled"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabIndex"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabPanelIndex"
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TabStop"
			Visible=true
			Group="Position"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Text"
			Visible=true
			Group="Initial State"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TextColor"
			Visible=true
			Group="Appearance"
			InitialValue="&h000000"
			Type="Color"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TextFont"
			Visible=true
			Group="Font"
			InitialValue="System"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TextSize"
			Visible=true
			Group="Font"
			InitialValue="0"
			Type="Single"
		#tag EndViewProperty
		#tag ViewProperty
			Name="TextUnit"
			Visible=true
			Group="Font"
			InitialValue="0"
			Type="FontUnits"
			EditorType="Enum"
			#tag EnumValues
				"0 - Default"
				"1 - Pixel"
				"2 - Point"
				"3 - Inch"
				"4 - Millimeter"
			#tag EndEnumValues
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Underline"
			Visible=true
			Group="Font"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UseFocusRing"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Visible"
			Visible=true
			Group="Appearance"
			InitialValue="True"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Width"
			Visible=true
			Group="Position"
			InitialValue="100"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
