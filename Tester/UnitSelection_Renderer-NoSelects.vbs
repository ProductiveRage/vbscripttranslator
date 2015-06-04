			'====================================================================================================
			' Control Overview
			'====================================================================================================
			' HaveInvalidOfferDetailsBeenProvided
			'  [Conf Booking validation, called from Render method]
			' RenderBookingInfoForm
			'  [Used by BookingUI_StayMain_AvailCal, BookingUI_StayMain_Legacy AND BookingUI_StayMain_Polling
			'   to render out common hidden form fields for each options form - only multiple if there is VB
			'   Polling or Fuzzy Searching enabled]
			'
			' BookingUI_StayMain_AvailCal
			' BookingUI_RenderNewReq_AvailCal
			'  [AvailCal rending uses these with RenderBookingInfoForm and BookingUI_RenderButtons only,
			'   BookingUI_StayMain_AvailCal is called direct from the Render method - Note: the Avail
			'   Cal format does not require any data access, so TOv2 is never accessed]
			'
			' BookingUI_StayMain
			'  [Calls BookingUI_StayMain_Legacy or BookingUI_StayMain_Polling]
			'
			' BookingUI_StayMain_Polling
			'
			' BookingUI_StayMain_Legacy
			'
			' BookingUI_StaySummary
			' BookingUI_StayDetails / BookingUI_StayDetails_PollingHeader / BookingUI_StayDetailsUrl
			' BookingUI_RenderNewStay / BookingUI_StayTtl / BookingUI_StayDiff
			' BookingUI_RenderNewReq
			' BookingUI_RenderUnit
			' BookingUI_RenderButtons
			' BookingUI_AvailClassName / BookingUI_AvailClassIcon
			' BookingUI_NicePrice
			' BookingUI_LinkedUnitDesc
			'
			' BookingUI_TicketsSummary
			'  [Tickets uses this to render content - called from BookingUI_StayMain_Legacy - doesn't need all
			'   of the above]
			'
			' GetExtBookUrlFromProductEstate
			' ====================================================================================================
			Option Explicit
			
			Dim Page, Request, Context, Server, DMS
			
			Dim InterfaceVersion

			InterfaceVersion = 1
			
			Const BOOKING_Local = 0
			Const BOOKING_Eviivo = 1
			Const BOOKING_Redirect = 2
			Const BOOKING_PollingRedirect = 3
			
			'nasty globals
			Dim IsExternalBooking, strExtBookUrl, strProductEstateID, bFormRendered, IsVBPollingEnabled, bRenderAsCalendar
			Dim g_iNumberOfCalendarsRendered : g_iNumberOfCalendarsRendered = 0
			bFormRendered = False

			Dim bProdHasAvail: bProdHasAvail = False
			Function GetProdHasAvail()
				GetProdHasAvail = bProdHasAvail
			End Function
			
			'public methods
			
			' ====================================================================================================
			' RENDER: Availability Calendar (supports local availability only!)
			' - Note: This doesn't actually perform any data access, all of the content required is passed
			'   through in POST data from the availability calendar on the previous page (Product Detail)
			' ====================================================================================================
			Function BookingUI_StayMain_AvailCal(ByVal pO, ByVal objRenderSettings)
				' Expect selections as set of form values:
				'  "unit_prodkey", "minoccu_prodkey", "maxcap_prodkey", "name_prodkey", "availclass_prodkey"
				'
				' If linked units are referenced, the first value will be:
				'  "unit_prodkey_linkprodkey"
				
				' 2011-08-09 DWR: Get populated read-only Booking Requirement data From GetSharedObject, then translate into a local copy we can edit 
				' (since some methods in here try to mess about with properties on it)
				Dim objBookingRequirement : Set objBookingRequirement = Page.Functions.GetSharedObject("BookingRequirement")
				Set objBookingRequirement = GetEditableBookingRequirement(objBookingRequirement)
				
				Dim intBookingType
				Dim iStayNum, iThisReqmnt
				Dim iUnitQty, iUnitMinOccupancy, iUnitMaxCapacity
				Dim iUnitKey, iLinkedUnitKey
				Dim strUnitName, strAvailClassId
				Dim Item, strTemp, i
				
				Dim dStart : dStart = objBookingRequirement.VisitDate
				Dim iNights : iNights = objBookingRequirement.Nights
				Dim intProdKey : intProdKey = objBookingRequirement.Product			
				
				' Open form and prepare to wrap content in "staySelection" container
				If (IsExternalBooking) Then
					intBookingType = BOOKING_Redirect
				Else
					intBookingType = BOOKING_Local
				End If
								
				RenderBookingInfoForm pO, intProdKey, objRenderSettings, intBookingType, Null, Null, Null, Null, Null, Null
				
				pO.Write "<div class=""staySelection"">"

				' Try to pull requirement info from Request
				iStayNum = 1
				iThisReqmnt = 0
				For Each Item In Request.Form
					'## Loop through only units
					If Left(Item, 5) = "unit_" Then

						strTemp = Right(Item, Len(Item) - 5)
						If Instr(strTemp, "_") > 0 Then
							' Linked unit
							iUnitKey = CLng(Right(strTemp, Len(strTemp) - Instr(strTemp, "_") ))
							iLinkedUnitKey = CLng(Left(strTemp,Instr(strTemp, "_")- 1))
						Else
							iUnitKey = CLng(strTemp)
							iLinkedUnitKey  = 0
						End If

						iUnitQty = CLng("0" & Request(Item))
						iUnitMinOccupancy = CLng("0" & Request("minoccu_" & strTemp))
						iUnitMaxCapacity = CLng("0" & Request("maxcap_" & strTemp))

						strUnitName = Request("name_" & strTemp)
						strAvailClassId = Request("availclass_" & strTemp)
						If iUnitQty > 0 Then
							For i = 1 To iUnitQty
								iThisReqmnt = iThisReqmnt+1
								BookingUI_RenderNewReq_AvailCal intBookingType, iUnitKey, strUnitName, iUnitMinOccupancy, iUnitMaxCapacity, strAvailClassId, iStayNum, iThisReqmnt, pO
							Next
							If iLinkedUnitKey > 0 Then
								pO.write "<input type=""hidden"" name=""linked_" & iUnitKey & """  value=""" & iLinkedUnitKey & """ />"
							End If
						End If
					End If
				Next
				
				' If successfully received requirement data, complete form - otherwise render error
				If iThisReqmnt > 0 Then
					pO.write "<input type=""hidden"" name=""availcal"" value=""" & Request("availcal") & """ />"
					pO.write "<input type=""hidden"" name=""_nStays"" value=""" & iStayNum & """ />"
					pO.write "<input type=""hidden"" name=""_nReqs"" value=""" & iThisReqmnt & """ />"

					' Close pnStayReqmntRslts div
					pO.write "</div>"

					BookingUI_RenderButtons iStayNum, pO, False

					' Close StayCandidateItem div
					pO.write "</div>"
				Else
					pO.Write Page.Resource("bookonline/unitselection/availcalendar/nounitsselectederror","<h2>Error</h2><p class=""error"">No units selected. Please click on the back button to return to the previous page and select the units you wish to book.</p>")
				End If

				' Close "staySelection" div and form
				pO.Write "</div>"
				pO.Write "</form>"
				If Page.Site.Params("Booking_ChildPricing") Then
					pO.Write "<script type=""text/javascript"">"
					pO.Write "NewMind.ETWP.Booking.UnitSelectionChildPricingGuests.Init();"
					po.Write "</" & "script>"
				End If
			End Function
			
			' ====================================================================================================
			' RENDER: Main entry point when not using Availability Calendar
			' ====================================================================================================
			' SUMMARY: entry point for product UNIT/STAY selection (Booking)
			' [aiProductKey]: integer product key
			' [adtStartNight]: date of first night of stay
			' [aiNights]: integer number of nights
			' [aiFuzzyStayNumDays]: integer flexible start date days (ZERO = Precise match)
			Function BookingUI_StayMain(objRenderSettings, objData)
				'most of these render functions rely on global variables, rather than trying to refactor them out for now ill create some globals
				'this needs refactoring
				
				' 2011-08-09 DWR: Expect the BookingRequirement in objRenderSettings to be read-only (since it usually comes from Page.Functions.GetSharedObject),
				' so replace it with an editable version (since some methods in here try to mess about with properties on it)
				Set objRenderSettings.BookingRequirement = GetEditableBookingRequirement(objRenderSettings.BookingRequirement)

				IsVBPollingEnabled = objRenderSettings.IsVBPollingEnabled
				bRenderAsCalendar = objRenderSettings.RenderAsCalendar
				If (objData Is Nothing) Then
					' If couldn't retrieve product, report no availability - this will happen if the
					' availability criteria can (no longer) be met
					RenderNoAvailElement objRenderSettings
					Exit Function
				End If

				If (objRenderSettings.LegacyRender) Then
					' Acco or Ticketing w/out VB Polling Enabled: Results from single Supplier (either
					' local OR FrontDesk for Acco, only local applies for Tickets)
					BookingUI_StayMain_Legacy objData, objRenderSettings
				Else
					' Acco w/ VB Polling Enabled: Results from multiple Suppliers
					' - Not supported when handling Conference Bookings, these are local only (but when
					'   an OfferKey is set, IsVBPollingEnabled is put to False - see PreRender)
					BookingUI_StayMain_Polling objData, objRenderSettings
				End If
			End Function
			
			'internal methods


			' ====================================================================================================
			' RENDER: Write out form with hidden input fields used for internal or FrontDesk bookings
			' - This will open the form, but the caller must close it
			' ====================================================================================================
			' Note: We need to pass intProdKey into here as we may not have an objProduct reference
			' (eg. if called by BookingUI_StayMain_AvailCal)
			Private Function RenderBookingInfoForm( _
				ByVal pO, _
				ByVal intProdKey, _
				ByVal objRenderSettings, _
				ByVal intBookingType, _
				ByVal strSupplierId, _
				ByVal strSupplierName, _
				ByVal strSupplierEviivoName, _
				ByVal strSupplierDeepLinkQuality, _
				ByVal strSupplierLogo, _
				ByVal intEviivoSearchIndustryClassification)

				' intBookingType can be one of:
				'  BOOKING_Local => Proceed to "Checkout" stage next
				'  BOOKING_Eviivo => Proceed to "Checkout" stage next, but handling a FrontDesk
				'  BOOKING_Redirect => Will redirect to complete booking on separate (probably NewMind) site
				'  BOOKING_PollingRedirect => Will proceed to "PollingExit" stage next
				' strSupplierDeepLinkQuality should be null unless intBookingType is BOOKING_PollingRedirect,
				' in which case it should be a string (possibly null if we didn't have this information available
				' in the DMS about the Supplier)
				Dim strPostUrl, strFormClass, strNextStage

				' Legacy rendering uses different form class for bookings that leave the site (we render them
				' all the same when VB Polling is enabled, though)
				If (intBookingType = BOOKING_Redirect) Then
					strFormClass = "FrmUnitOptionsExt"
				Else
					strFormClass = "FrmUnitOptions"
				End If

				' What booking stage is next?
				' - If not external, go to checkout regardless of VB Polling setting.
				' - If IS external, branch off differently (VB Polling goes to a separate switcher stage, non-
				'   VB-Polling will redirect to the other site).
				' While we're here, retrieve POST url (secure for checkout, standard otherwise)
				If ((intBookingType = BOOKING_Local) Or (intBookingType = BOOKING_Eviivo)) Then
					strNextStage = "checkout"
					strPostUrl = GetPostUrl(True) & "/" & strNextStage
				ElseIf (intBookingType = BOOKING_Redirect) Then
					'strNextStage = "redirect"
					strNextStage = "checkout" 'This should stay as "checkout" until 1.4 is updated to recognise "redirect" stage
					strPostUrl = Page.PageInfo.GetUrlFromPageID("EXTBOOKPROMPT")
					If IsNull(strPostUrl) Then
						Page.PrintTraceWarning "RenderBookingInfoForm: Unable to locate page EXTBOOKPROMPT, default to current page - is this correct behaviour??"
						strPostUrl = Page.URL.Real
					End If
				ElseIf (intBookingType = BOOKING_PollingRedirect) Then
					' 2014-06-19 DWR: We have historically used the SupplierEviivoName for the URL segment, although it used to be labelled strSupplierName since
					' the values were getting set incorrectly. SupplierEviivoName seems like the most appropriate option since it will be a text-friendly string
					' value and so not have dots or spaces or whatever (and so be good for use in a URL).
					strNextStage = "pollingexit"
					strPostUrl = GetPostUrl(False) & "/pollingexit/" & strSupplierEviivoName
				Else
					Err.Raise vbObjectError, "ETWP.BookingUnitSelection", "RenderBookingInfoForm: Invalid intBookingType value (" & intBookingType & ")"
				End If

				pO.Write "<form action=""" & strPostUrl & """ "
				If (Not IsVBPollingEnabled) And objRenderSettings.BookingRequirement.FlexibleRange = 0 Then
					' Can't have ids when VB Polling enabled as we might be rendering out multiple of these forms.
					' 2008-11-10 DWR: This is similarly the case for fuzzy searching. I don't we have any working
					' Enterprise fuzzy-searching sites, so don't need to worry about breaking styling by removing
					' this id in this case.
					pO.Write "id=""FrmUnitOptions"" "
				End If

				'#MJ's Reasoning -	In order for us to jump to unit selection in a tab it must have a name, however only the first form should have this
				If Not bFormRendered Then
					pO.Write "name=""FrmUnitOptions"" "
					bFormRendered = True
				End If
				pO.Write "class=""" & strFormClass & """ method=""post"">"

				' Open container around common form values
				pO.Write "<div>"

				pO.Write "<input type=""hidden"" name=""stage"" value=""" & strNextStage & """ />"

				' Need to override market source if viewing site via widget
				If Page.WidgetView Then
					If (intBookingType = BOOKING_Redirect) then
						' External bookings visit a preliminary redirect page first, which we want to be decluttered when in a widget
						pO.Write "<input type=""hidden"" name=""widget_marketsource"" value=""" & Page.WidgetMarketSource & """ />"
					Else
						pO.Write "<input type=""hidden"" name=""msource"" value=""" & Page.WidgetMarketSource & """ />"
					End If
					'this hidden field is to tell the checkout that weve come from a widget, and not a failed checkout validation
					pO.Write "<input type=""hidden"" name=""widget"" value=""1"" />"
				End If

				' None of this applies to VB Polling, even if it IS an external booking - we go to an
				' interim stage before leaving the site
				If (intBookingType = BOOKING_Redirect) Then
					' NB: In "Conference Booking" mode (where OfferKey <> 0), we need to set the "channel" and "msource"
					'     values to different values (for msource, if there is no "ConfBookingMarketSourceID" set, it will
					'     fall back to using the site's main "MarketSourceID" source)
					pO.Write "<input type=""hidden"" name=""checkoutstage"" value=""1"" />"
					If (objRenderSettings.BookingRequirement.Offer = 0) Then
						pO.Write "<input type=""hidden"" name=""channel"" value=""" & objRenderSettings.Channel & """ />"
					Else
						pO.Write "<input type=""hidden"" name=""channel"" value=""" & objRenderSettings.ConfBookingChannel & """ />"
					End If
					If Not Page.WidgetView Then
						'Neeed to set market source override if redirecting to external site unless set above due to widgetview
						If (objRenderSettings.BookingRequirement.Offer = 0) Or (Page.Site.Params("ConfBookingMarketSourceID") = "") Then
							pO.Write "<input type=""hidden"" name=""msource"" value=""" & Page.Site.Params("MarketSourceID") & """ />"
						Else
							pO.Write "<input type=""hidden"" name=""msource"" value=""" & Page.Site.Params("ConfBookingMarketSourceID") & """ />"
						End If
					End If
					pO.Write "<input type=""hidden"" name=""bookchannel"" value=""" & Page.Site.Params("Booking_ChannelID") & """ />"
					pO.Write "<input type=""hidden"" name=""reposturl"" value=""" & strExtBookUrl & """ />"
					' 2009-09-21 DWR: New field to pass in so that the receiving site recognises booking as having
					' come from another site (so it can update appropriate Provider Stats)
					pO.Write "<input type=""hidden"" name=""ForcedExternalBooking"" value=""1"" />"
				End If

				pO.Write "<input type=""hidden"" name=""product"" value=""" & intProdKey & """ />"
				pO.Write "<input type=""hidden"" name=""isostartdate"" value=""" & Page.Functions.Dates.ISODate(objRenderSettings.BookingRequirement.VisitDate) & """ />"
				pO.Write "<input type=""hidden"" name=""nights"" value=""" & objRenderSettings.BookingRequirement.Nights & """ />"

				' We need all this when using VB Polling, even it it is an external booking, as we aren't
				' going to leave the site yet (there's an interim stage)
				If (intBookingType <> BOOKING_Redirect) Then
						' NB: "package" parameter removed - it's now passed as "offer", and only when
						' customer is going for a "Conference Booking" discount product.
					' 2008-11-07 DWR: This used to referer to a "strRewriteUrl" value that was never defined.
					' So we'll pass in blank. Pretty sure it's not used anyway.
					pO.Write "<input type=""hidden"" name=""preUrl"" value="""" />"
					' 2008-11-07 DWR: If we've got non-precise results from a fuzzy search, we'll render this
					' form out and use the actual StartDate / NumNights combination that the fuzzy results
					' offered. So we just pass these to the checkout stage, and set "fuzzy" to zero.
					pO.Write "<input type=""hidden"" name=""fuzzy"" value=""0"" />"
					pO.Write "<input type=""hidden"" name=""lng"" value="""&Page.Language.LanguageCultureKey&""" />"

					' NB: OfferKey is required for products in the "Conference Booking" functionality as
					' it lets the checkout object know that we should be looking for the product on the
					' "Conference Booking Channel" instead of the standard "website" channel. If this
					' ever needed to work with the ExternalBooking, we would need to pass out the
					' conference channel in the IsExternalBooking section above, but since this is
					' only being supported by the internal Newmind booking, it's not an issue.
					pO.Write "<input type=""hidden"" name=""offer"" value=""" & objRenderSettings.BookingRequirement.Offer & """ />"

					' Pass in the current convert-to-currency value (this will have been held in the session
					' up to this point, but we may be about to leave the site when this form is posted, so
					' will need to send the value as a hidden input instead of relying on session)
					pO.Write "<input type=""hidden"" name=""CurrencyConvertTo"" value=""" & Page.Functions.Money.GetCurrencyCodeOverride(Page.Site.LCCurrencyKey) & """ />"
				End If

				' If we're dealing with a VB Polling External Supplier, write out the Supplier id, name and
				' deep-link-quality as well (this is the number of rooms that the supplier can handle in
				' deep-linking situations)
				If (intBookingType = BOOKING_PollingRedirect) Then
					pO.Write "<input type=""hidden"" name=""SupplierId"" value=""" & strSupplierId& """ />"
					pO.Write "<input type=""hidden"" name=""SupplierName"" value=""" & strSupplierName & """ />"
					pO.Write "<input type=""hidden"" name=""SupplierLogo"" value=""" & strSupplierLogo & """ />"
					pO.Write "<input type=""hidden"" name=""SupplierEviivoName"" value=""" & strSupplierEviivoName & """ />"

					pO.Write "<input type=""hidden"" name=""EviivoSearchIndustryClassification"" value="""
					If IsNumeric(intEviivoSearchIndustryClassification) Then
						pO.Write intEviivoSearchIndustryClassification
					Else
						pO.Write "0"
					End If
					pO.Write """ />"
					
					If IsNull(strSupplierDeepLinkQuality) Then
						strSupplierDeepLinkQuality = ""
					Else
						strSupplierDeepLinkQuality = Trim(strSupplierDeepLinkQuality)
					End If
					If Not IsNumeric(strSupplierDeepLinkQuality) Then
						strSupplierDeepLinkQuality = "-1"
					End If
					pO.Write "<input type=""hidden"" name=""SupplierDeepLinkQuality"" value=""" & strSupplierDeepLinkQuality & """ />"
				End If

				' Append in the "Nominal Units" from Request collection or objUnitReqDictFromBookUrl (ie. "roomReq_1", "roomReq_2", etc..)
				'#MJ TODO need to call the new function
				pO.Write GenerateRequirementFormData(objRenderSettings.BookingRequirement)
				' Close common form value container
				pO.Write "</div>"
				
			End Function

			'generates a string of room requirement details in a format suitable for use in a form i.e hidden inputs ;)
			Private Function GenerateRequirementFormData(ByRef objAccoSearchRequirement)
				'get our key value data dictionary
				Dim dictKeyValues : Set dictKeyValues = Page.Functions.Booking.GenerateRequirementKeyValueData(objAccoSearchRequirement)
				'create an array to hold our formatted data in which is the same size of the dictionary
				Dim aryFormattedData
				ReDim aryFormattedData(dictKeyValues.Count - 1)
				'spin through our output array and add the formatted items in the format {key}={value}
				Dim i : i = 0
				Dim key
				For Each key in dictKeyValues.Keys
					'#MJ's Reasoning -	we don't want to render our roomrequirement's here as they may not be valid
					'					instead when we write out a room requirement say what room it is and for how many
					' NP: MJ is saying add hidden form values based on the requirements linked to the UnitStayDetails by the AvailCom 
					' NOT to base it on the BookingRequestDictionary. These form values will then be posted 
					' and update the BookingRequirement object for when it is used in the Booking Checkout
					If Not Left(LCase(key), 8) = "roomreq_" Then
						aryFormattedData(i) =  "<input type=""hidden"" name="""&key&""" value="""&dictKeyValues.Item(key)&""" />"&vbCRLF
					End If
					i = i + 1
				Next
				'return our array as a string using an & as the joining character
				GenerateRequirementFormData = Join(aryFormattedData)
			End Function

			Private Function GetPostUrl(ByVal bSecure)
				Dim strPostUrl, strUrl
				
				If (bSecure) Then
					strPostUrl = Page.Site.SecureHostName
				Else
					strPostUrl = Page.URL.FullHostName
				End If
				Do While (Right(strPostUrl, 1) = "/")
					strPostUrl = Left(strPostUrl, Len(strPostUrl) - 1)
				Loop

				strUrl = Page.PageInfo.GetUrlFromPageID("BOOKONLINE")
				If IsNull(strUrl) Then
					Page.PrintTraceWarning "GetPostUrl: Unable to locate page BOOKONLINE, default to current page - is this correct behaviour??"
					strURL = Page.URL.Real
				End If

				If (Left(strUrl, 1) <> "/") Then
					strUrl = "/" & strUrl
				End If

				strPostUrl = strPostUrl & strUrl
				If (UCase(Left(strPostUrl, 7)) <> "HTTP://") And (UCase(Left(strPostUrl, 8)) <> "HTTPS://") Then
					strPostUrl = "http://" & strPostUrl
				End If
				
				GetPostUrl = strPostUrl
			End Function

			

			' SUMMARY: render new requirement UI from avail calendar
			' [ireqSz]: ADO unit recordset from availability object
			' [aiStayNum]: integer stay index
			' [aiThisReqmnt]: integer requirement number (from recordset)
			Private Function BookingUI_RenderNewReq_AvailCal(intBookingType, iUnitKey, strUnitName, iUnitMinOccupancy, iUnitMaxCapacity, asAvailClassId, aiStayNum, aiThisReqmnt, pO)
			' on first ever call [aiThisReqmnt]=1, on subsequent calls we must close previous [pnStayReqmnt] and [pnStayReqmntRslts] DIVs
				Dim iGuest
				
				If aiThisReqmnt>1 Then
					pO.Write "</div></div>"
				End If

				pO.Write "<div class=""pnStayReqmnt"">"&vbCrLf
				pO.Write "<div class=""pnStayReqmntTtl"">"&vbCrLf
				pO.Write "<div Class=""pnStayReqmntRoom"">Room "&aiThisReqmnt&" - "&strUnitName&BookingUI_AvailClassIcon(asAvailClassId)&" <br/></div>"

				If iUnitMinOccupancy = 0 Or iUnitMinOccupancy = ""  Then
					iUnitMinOccupancy = 1
				End If

				pO.Write "<div Class=""pnStayReqmntGuests"">"&vbCrLf
				If iUnitMaxCapacity = iUnitMinOccupancy Then
					pO.Write "For "&iUnitMaxCapacity&" guests <input type=""hidden"" name=""roomReq_"&aiThisReqmnt&""" value="""&iUnitMaxCapacity&"""/>"
				Else
					Dim strGuestsFor : strGuestsFor = Page.Resource("bookonline/unitselection/guestrequirement/for", "for") 
					'alas child pricing is different
					If Page.Site.Params("Booking_ChildPricing") Then
					
						Dim strAdultsTitle : strAdultsTitle = Page.Resource("bookonline/unitselection/guestrequirement/adults/selecttitle", "Please specify the number of adults in this room.")
						Dim strAdults : strAdults = Page.Resource("bookonline/unitselection/guestrequirement/adults/adult(s)", "adult(s)")
						
						pO.Write strGuestsFor & " <select class=""adults"" name=""roomReq_"&aiThisReqmnt&"_adults"" title="""&strAdultsTitle&"""> "
						For iGuest = iUnitMinOccupancy To iUnitMaxCapacity
							pO.Write "<option value=""" & iGuest & """>" & iGuest & "</option> "
						Next
						pO.Write "</select> " & strAdults
						
						Dim strChildrenTitle : strChildrenTitle = Page.Resource("bookonline/unitselection/guestrequirement/children/selecttitle", "Please specify the number of children in this room.")
						Dim strChildren : strChildren = Page.Resource("bookonline/unitselection/guestrequirement/children/children", "children")
						Dim strGuestsAnd : strGuestsAnd = Page.Resource("and", "and")
						pO.Write " " & strGuestsAnd & " <select class=""children"" name=""roomReq_"&aiThisReqmnt&"_children"" title="""&strChildrenTitle&"""> "
						
						For iGuest = 0 To iUnitMaxCapacity - 1
							pO.Write "<option value=""" & iGuest & """>" & iGuest & "</option> "
						Next
						pO.Write "</select> " & strChildren
						
						Dim iCount
						pO.WriteLine "<span class=""label childrenageslabel"">Child Ages</span>"
						pO.WriteLine "<span class=""field childrenagesfield"">" 
						
						Dim ageValue
						For iCount = 0 To (iUnitMaxCapacity - 1)
							pO.WriteLine "<span class=""childageWrapper"">"
							pO.WriteLine vbTab & "<span class=""label childagelabel"">Child Age " &  iCount + 1 & "</span>"
							pO.WriteLine vbTab & "<span class=""field childagefield"">"
							pO.Write "<select class="""" name=""roomReq_"&aiThisReqmnt&"_children_childage"&iCount&""">"
							For iGuest = 0 To 18
								pO.Write "<option value=""" & iGuest & """>" & iGuest & "</option> "
							Next
							pO.Write "</select> "
							pO.WriteLine vbTab & "</span>"
							pO.WriteLine "</span>"
						Next
						pO.WriteLine "</span>"
					Else
						Dim strGuestsTitle : strGuestsTitle = Page.Resource("bookonline/unitselection/guestrequirement/selecttitle", "Please specify the number of guests in this room.")
						Dim strGuests : strGuests = Page.Resource("bookonline/unitselection/guestrequirement/guest(s)", "guest(s)")
						
						pO.Write strGuestsFor & " <select name=""roomReq_"&aiThisReqmnt&""" title="""&strGuestsTitle&"""> "
						For iGuest = iUnitMinOccupancy To iUnitMaxCapacity
							pO.Write "<option value=""" & iGuest & """>" & iGuest & "</option> "
						Next
						pO.Write "</select> " & strGuests
					End If
				End If
				pO.Write "</div>"&vbCrLf

				If intBookingType = "ticketing" Then
					pO.Write "<input type=""hidden"" name=""unit_"&iUnitKey&"""  value="""&aiThisReqmnt&""" />"
				Else
					pO.Write "<input type=""hidden"" name=""unit_"&aiStayNum&"_"&aiThisReqmnt&"""  value="""&iUnitKey&""" />"
				End If
			End Function

			

			' SUMMARY: Draw availability month calendar
			' [sbCalendars]:  ASP [nmStringBuilder] object instance output string
			' [dCalStartDflt]: date default calendar start date
			' <retval>: string month available stays details JSON data
			Private Function BookingUI_RenderAvailCal(ByRef sbCalendars, objDictAvaiStays, bStarted)	

				Dim strClassMonth : strClassMonth = "MonthWrapper"
				If Not bStarted Then
					sbCalendars.AppendLine("<div class=""CalendarsWrapper"">") 
					sbCalendars.AppendLine("<div class=""instruction"">"&Page.Resource("bookonline/unitselection/availcalendar/instruction","Please select an available stay from the calendars below. Clicking on a highlighted start day for a stay will show the stay details such as the units available, price, etc.") & "</div>")
					strClassMonth = strClassMonth & " currentmonth"
				Else 
					strClassMonth = strClassMonth & " nextmonth"
				End If 
				
				Dim dStart1, aryAvailStaysKeys
				If objDictAvaiStays.Count > 0 Then 
					aryAvailStaysKeys = objDictAvaiStays.Keys
					dStart1 = Replace(aryAvailStaysKeys(0), "sd_", "")
					Erase aryAvailStaysKeys
				Else
					dStart1 = Date()
				End If
				
				BookingUI_RenderCalendarMonthWithAvailability sbCalendars, dStart1, strClassMonth, objDictAvaiStays

				' using a global count so we can track how many calendars have been added to the stringbuilder for the prev/next buttons
				' doing this now because of the recursive nature of this function
				g_iNumberOfCalendarsRendered = g_iNumberOfCalendarsRendered + 1		
					
				'Check if we have stays left and render then as another calendar
				If objDictAvaiStays.Count > 0 Then
					BookingUI_RenderAvailCal sbCalendars, objDictAvaiStays, True
				Else
					'not sure if this should be dStart1 - was dStart
					BookingUI_RenderAvailCalLinks dStart1, sbCalendars
					BookingUI_RenderAvailCalKey sbCalendars
					sbCalendars.AppendLine("</div>")		
				
				End If 
			
			End Function  

			Private Function BookingUI_RenderCalendarMonth(ByRef sbCalendars, ByVal dFirstDayOfMonth, ByVal strWrapperClass)
				BookingUI_RenderCalendarMonthWithAvailability sbCalendars, dFirstDayOfMonth, strWrapperClass, Nothing
			End Function

			Private Function BookingUI_RenderCalendarMonthWithAvailability(ByRef sbCalendars, ByVal dFirstDayOfMonth, ByVal strWrapperClass, ByVal objDictAvailStays)
				
				Dim iWeekStartDay : iWeekStartDay = 1 'Monday		
				Dim iWeekDayCalStart : iWeekDayCalStart = (iWeekStartDay + 1) Mod 7
				Dim iWeekDayCalEnd : iWeekDayCalEnd =  iWeekStartDay Mod 7

				Dim dCalStart : dCalStart = Page.Functions.Dates.fn_GetFirstDateOfMonth(dFirstDayOfMonth)
				Dim dCalEnd : dCalEnd = Page.Functions.Dates.fn_GetLastDateOfMonth(dFirstDayOfMonth)
				Dim strThisMonthYear : strThisMonthYear = Page.Functions.Dates.GetMonthNameAbbr(Month(dCalStart)) & " " & Year(dCalStart) 
				Dim strTableSummary : strTableSummary = Page.Resource("bookonline/unitselection/availcalendar/availabilitycalendarfor","Availability calendar for") & " " & strThisMonthYear

				sbCalendars.AppendLine("<div id=""Cal_" & Page.Functions.Dates.ISODate(dCalStart) & """ class=""" & strWrapperClass & """>")
				sbCalendars.AppendLine("<table id=""Tbl_" & Page.Functions.Dates.ISODate(dCalStart) & """ class=""availabilityCalendar"" summary=""" & strTableSummary & """ >")
				sbCalendars.AppendLine("<thead>")
				sbCalendars.AppendLine("<tr>")
				sbCalendars.AppendLine("<th colspan=""8"">" & strThisMonthYear & "</th>")
				sbCalendars.AppendLine("</tr>")

				Dim strHeaderCellClass : strHeaderCellClass = ""
				Dim i : For i = iWeekStartDay to iWeekStartDay + 6
					If (i Mod 7) = 6 or (i Mod 7) = 0 Then
						strHeaderCellClass = " class=""we"""
					End If
					sbCalendars.AppendLine("<th" & strHeaderCellClass & ">" & Page.Functions.Dates.GetDayNameAbbr(Weekday((i + 1) Mod 7)) & "</th>")
				Next
			
				sbCalendars.AppendLine("</tr>")
				sbCalendars.AppendLine("</thead>")
				sbCalendars.AppendLine("<tbody>")
			
				Dim iCellCount : iCellCount = 0
				Dim bFirstCell : bFirstCell = True
				Dim bLastCell : bLastCell = False
				
				Dim dDate : dDate = dCalStart		

				Dim bStartNewStay, bStayIndicative
				Dim strStayNumber

				Dim iDay : For iDay = Day(dCalStart) To Day(dCalEnd)
					bStartNewStay = False
			
					If bFirstCell Then
						Dim iPrePadding : iPrePadding = DateDiff("d", Page.Functions.Dates.fn_GetFirstDateOfWeek(Page.Functions.Dates.fn_GetFirstDateOfMonth(dCalStart), iWeekDayCalStart), dCalStart)
						If iPrePadding > 0 Then
							Dim j : For j = 1 To iPrePadding
								sbCalendars.AppendLine("<td></td>")
								iCellCount = iCellCount + 1
							Next
						End If
						bFirstCell = False
					End If
					
					Dim strDisplayText : strDisplayText = "" & iDay
					Dim strDayCellClass : strDayCellClass = "n"
		
					If Not objDictAvailStays Is Nothing Then 
						If objDictAvailStays.Exists("sd_" & dDate) Then 
							bStartNewStay = True 
							'we expect value in the format [stayNo]_[indicative]
							Dim aryStay : aryStay = Split(objDictAvailStays("sd_" & dDate), "_")
							strStayNumber = aryStay(0)
							bStayIndicative = CBool(aryStay(1))
							objDictAvailStays.Remove("sd_" & dDate)
							Erase aryStay
						End If 
					End If 

					If dDate < Date() Then 'date is in the past
					
						strDayCellClass = "p"	
				
					ElseIf bStartNewStay Then 
					
						If Not objDictAvailStays Is Nothing Then 
							Dim strAvailType : strAvailType = ""
							Dim strIndicativeIcon : strIndicativeIcon = ""

							If bStayIndicative Then
								strDayCellClass = "i"
								strAvailType = Page.Resource("bookonline/unitselection/unconfirmedavailability","Unconfirmed Availability")
								strIndicativeIcon = "<img src=""" & Page.ImageResource("bookonline/icons/indicative", "/images/icon_indicative.gif") & """ alt=""" & strAvailType & """ class=""icon""/>"
							Else
								strDayCellClass = "a"
								strAvailType = Page.Resource("bookonline/unitselection/confirmedavailability","Confirmed Availability")
								strIndicativeIcon = "<img src=""" & Page.ImageResource("bookonline/icons/allocated", "/images/icon_allocated.gif") & """ alt=""" & strAvailType & """ class=""icon""/>"
							End If
								
							strDisplayText = "<a href=""#stay_" & strStayNumber & """ class=""calavailstay"" id=""stay_" & strStayNumber & """>" & Day(dDate) & "</a>" & strIndicativeIcon						
						
						End If	
					
					End If 				
		
					If Weekday(dDate) = 1 or Weekday(dDate) = 7 Then
						strDayCellClass = strDayCellClass & " we"
					End If
														
					sbCalendars.AppendLine("<td class=""" & strDayCellClass & """><div>" & strDisplayText & "</div></td>")
					
					iCellCount = iCellCount + 1

					If dDate = dCalEnd Then
						bLastCell = True
					End If

					' This is for when the last day of the month is not the last day of the week and empty cells are put in place to fill the calendar days
					If bLastCell Then
						Dim iPostPadding : iPostPadding = DateDiff("d", dCalEnd, Page.Functions.Dates.fn_GetLastDateOfWeek(dCalEnd, iWeekDayCalEnd))
						If iPostPadding > 0 And iPostPadding < 7 Then
							Dim k : For k = 1 To iPostPadding
								sbCalendars.AppendLine("<td></td>")
								iCellCount = iCellCount + 1
							Next
						End If
						bLastCell = False
						bFirstCell = True
					End If
				
					If iCellCount Mod 7 = 0 Then
						sbCalendars.AppendLine("</tr>")
					End If
				
					dDate = DateAdd("d", 1, dDate)
				Next 
	
				sbCalendars.AppendLine("</tbody>")
				sbCalendars.AppendLine("</table>")
				sbCalendars.AppendLine("</div>")
			
			End Function
			
			Private Function BookingUI_RenderAvailCalKey(sb)
				Dim strCalKey : strCalKey = Page.Resource("bookonline/unitselection/availcalendar/calkey","")
				If Trim(strCalKey) <> "" Then 
					sb.AppendLine("<div class=""CalKey"">" & strCalKey & "</div>")
				End If 
			End Function
			
			Private Function BookingUI_RenderAvailCalLinks(dStart,sb)
				
				Dim dCalStartPrev,strTitlePrev, dCalStartNext, strTitleNext
				
				' dStart is the start date for the last month shown in the rendered calendars
				' and we therefore only need to go forward by 1 month 
				' even if no calendars are shown for the current month we can still potentially
				' move to a future month where there is availability.
				Dim iPositiveMonthAdjustment : iPositiveMonthAdjustment = 1
				' The previous month link has to go back by however many months are already showing, i.e. Jul & Aug are shown
				' dStart = 01/08/2011 (Aug) and we need to display Jun & Jul so we need to jump back 2 months to June. 
				Dim iNegativeMonthAdjustment : iNegativeMonthAdjustment = -(g_iNumberOfCalendarsRendered)

				If g_iNumberOfCalendarsRendered = 0 Then
					' If we have no rendered calendars we still need the link to go back by 1 month
					iNegativeMonthAdjustment = -1
				End If

				dCalStartPrev = Page.Functions.Dates.fn_GetFirstDateOfMonth(DateAdd("m", iNegativeMonthAdjustment, dStart))
				strTitlePrev = Page.Resource("bookonline/unitselection/availcalendar/previousmonth","&lt;&lt; Previous Month")
				
				dCalStartNext = Page.Functions.Dates.fn_GetFirstDateOfMonth(DateAdd("m", iPositiveMonthAdjustment, dStart))
				strTitleNext = Page.Resource("bookonline/unitselection/availcalendar/nextmonth","Next Month &gt;&gt;")
					
				sb.AppendLine "<div class=""CalNavLinks"">"
				sb.AppendLine BookingUI_RenderAvailCalLink(dCalStartPrev, strTitlePrev, "prev")
				sb.AppendLine BookingUI_RenderAvailCalLink(dCalStartNext, strTitleNext, "next")
				sb.AppendLine "</div>"
			
			End Function 
			
			Private Function BookingUI_RenderAvailCalLink(dCalStartDate, strTitle, strClass)
			
				Dim itm, sValue, strLink, bFound
				
				bFound = False
				
				For Each itm In Request.QueryString
					If itm="isostartdate" Then
						'reset date
						sValue = Page.Functions.Dates.ISODate(dCalStartDate)
						bFound = True 
					Else
						sValue = Request.QueryString(itm)
					End If 
					strLink = strLink &"&amp;"& itm & "=" &Server.UrlEncode(sValue)
				Next 
			
				If Not bFound Then strLink = strLink &"&amp;isostartdate=" &Server.UrlEncode(Page.Functions.Dates.ISODate(dCalStartDate))
				
				If Trim(""&strLink) <> "" Then 	strLink = Replace(strLink,"&amp;", "?", 1,1,0)
			
				If DateDiff("m",Date(),dCalStartDate) >= 0 Then 
					BookingUI_RenderAvailCalLink = "<a href="""&strLink&""" class="""&strClass&""" title="""&strTitle&""" rel=""nofollow"">"&strTitle&"</a>" & vbCrLf
				Else
					BookingUI_RenderAvailCalLink = ""
				End  If 
				
			End Function 
			

			

			' ====================================================================================================
			' RENDER: Main entry point when VB Polling is enabled
			' - Applies to acco products only
			' - Not supported when handling Conference Bookings (these are local only)
			' ====================================================================================================
			Private Function BookingUI_StayMain_Polling(objData, objRenderSettings)
				
				Dim pO : Set pO = objRenderSettings.OutputWriter
				Dim dStartNight : dStartNight = objRenderSettings.BookingRequirement.VisitDate
				Dim iNights : iNights = objRenderSettings.BookingRequirement.Nights

				' This is new, VB Polling approach (only supports accommodation, but handles results from
				' multiple providers)
				Dim objAvail: Set objAvail = objData.Availability
				Dim intProdKey: intProdKey = objData.Product_Key
				Dim bIsTeleBooking: bIsTeleBooking = objData.IsOnTeleBookingChannel
				
				Dim objAvailEntry
				Dim bNoResults
				Dim bRenderedSummary
				Dim intIndex, intIndexSupplier
				Dim objFuzzyStayOptions, objFuzzyStay, bPreciseMatch, bStayHasLocalAvail
				Dim objSuppliersForStay, objSupplier

				Dim objDictAvaiStays, strAvailStayKey, aryStay, sStayNo, bStayIndicative
				Set objDictAvaiStays = Server.CreateObject("Scripting.Dictionary")		

				' Quick situation assertion
				If (objRenderSettings.BookingType <> "accommodation") Then
					Err.Raise vbObjectError, "ETWP.BookingUnitSelection", "BookingUI_StayMain_Polling: BookingType not supported (""" & BookingType & """)"
				End If
				If (objRenderSettings.BookingRequirement.Offer > 0) Then
					Err.Raise vbObjectError, "ETWP.BookingUnitSelection", "BookingUI_StayMain_Polling: Not supported with Conference Bookings (OfferKey = " & objRenderSettings.BookingRequirement.Offer & ")"
				End If
				
				' Grab hold of the data for the stay(s) - ensure we've got some availability
				Set objFuzzyStayOptions = objAvail.GetUniqueFuzzyCombinations()
				If (objFuzzyStayOptions.Count = 0) Then
					Page.PrintTraceWarning "objAvail.GetUniqueFuzzyCombinations reported zero stay options"
					bNoResults = True
				Else
					' Double-check that all stay options report availability - there shouldn't be any stay
					' data returned that doesn't have avail data
					
					bNoResults = False
					For intIndex = 0 To objFuzzyStayOptions.Count - 1
						Set objFuzzyStay = objFuzzyStayOptions.GetItem(intIndex)
						Set objSuppliersForStay = objAvail.GetSupplierUnitDataForStay(objFuzzyStay.StartDate, objFuzzyStay.Nights)
						If (objSuppliersForStay.Count = 0) Then
							Page.PrintTraceWarning "Stay (" & objFuzzyStay.StartDate & ", " & objFuzzyStay.Nights & ") reported zero suppliers"
							bNoResults = True
						Else
							For intIndexSupplier = 0 To objSuppliersForStay.Count - 1
								Set objSupplier = objSuppliersForStay.GetItem(intIndexSupplier)
								If (objSupplier.Units.Count = 0) Then
									Page.PrintTraceWarning "Supplier " & objSupplier.Name & " for Stay (" & objFuzzyStay.StartDate & ", " & objFuzzyStay.Nights & ") reported zero units"
									bNoResults = True
								End If
							Next
						End If
					Next
				End If
						
				' If not, render error and get out
				If (bNoResults) Then
					' Render message, set ProdHasAvail To False (only
					' used by BookingKeys control, I think) and close recordsets
					RenderNoAvailElement objRenderSettings
					bProdHasAvail = False ' This is exposed through the WSC's public property "ProdHasAvail"
					Exit Function
				End If
				
				bProdHasAvail = True ' This is exposed through the WSC's public property "ProdHasAvail"
				
				' Loop through different stay options
				' - Store data for all stays for calendar
				If bRenderAsCalendar Then 

					For intIndex = 0 To objFuzzyStayOptions.Count - 1
						Set objFuzzyStay = objFuzzyStayOptions.GetItem(intIndex)

						strAvailStayKey = "sd_" & objFuzzyStay.StartDate
						If objDictAvaiStays.Exists(strAvailStayKey) Then 
		
							' We expect value in the format [stayNo]_[indicative]
							aryStay = Split(objDictAvaiStays(strAvailStayKey), "_")
							sStayNo = aryStay(0)
							sStayNo = sStayNo & "-" &intIndex
						
							bStayIndicative = CBool(aryStay(1))
							If Not bStayIndicative And objFuzzyStay.Indicative Then
								bStayIndicative = objFuzzyStay.Indicative
							End If
							objDictAvaiStays(strAvailStayKey) = sStayNo & "_" & bStayIndicative 
							Erase aryStay
						Else 
							objDictAvaiStays.Add "sd_" & objFuzzyStay.StartDate, intIndex + 1 & "_" & objFuzzyStay.Indicative
						End If
						
					Next

				End If

				Dim bRenderedInitialStay : bRenderedInitialStay = False
				' - For unit selections: If we have a perfect match stay, don't bother with the fuzzy options
				For intIndex = 0 To objFuzzyStayOptions.Count - 1
					' 2010-11-03 TB: Stay numbers are 1-based so add 1 to zero-based index
					Dim iStayNum : iStayNum = intIndex + 1
					
					Set objFuzzyStay = objFuzzyStayOptions.GetItem(intIndex)
					
					' 2010-01-29 DWR: Need to use DateValue here since dStartNight might be a string
					' which will cause the comparison to fail when they represent the same date
					bPreciseMatch = (DateValue(objFuzzyStay.StartDate) = DateValue(dStartNight)) And (objFuzzyStay.Nights = iNights)
					
					' 2010-01-29 DWR: In cases where we have a precise match and we're not rendering a calendar
					' then we want to just display that stay and get out! If we DON'T have a precise match and
					' we're not using the calendar approach then we want to render all options and have client
					' side script juggle them. If we ARE rendering the calendar then we want to display ALL
					' stays - regardless of whether we have a precise match - because the calendar relies
					' on the data being in the markup for it to swap around.
					If (bRenderAsCalendar) Or (Not bPreciseMatch) Then
						pO.Write "<div class=""PollingFuzzySetWrapper"" id=""stay_" & iStayNum & """>" 
					End If
					
					' we only render the first stay when initially loading the unitselection
					' we get the rest via a partial render request and the data is returned as JSON
					' ready for manipulation and insertion by javascript
					' this is done to avoid large amounts of HTML being rendered and then hidden
					If Not (bRenderedInitialStay) Then
						RenderStay objFuzzyStay, objAvail, iStayNum, objRenderSettings, bIsTeleBooking, objData.bookingweb, objData.EviivoId, objData.Units
					End If
					
					If (bRenderAsCalendar) Then
						bRenderedInitialStay = True
					End If
					
					' Close the wrapper for the current stay date/length result set
					' 2010-01-29: See earlier comment about this..
					If (bRenderAsCalendar) Or (Not bPreciseMatch) Then
						pO.Write "</div>"
					End If
					
					' If these options were a perfect match, drop out
					' 2010-01-29 DWR: Unless we're rendering the calendar! In this case client-side javascript
					' will look after showing one fuzzy stay at a time, but it needs all data present.
					If (bPreciseMatch) And (Not bRenderAsCalendar) Then
						Exit For
					End If
					
				Next

				If bRenderAsCalendar Then
						
					Dim ReqDictTemp: Set ReqDictTemp = Page.Functions.GetNewObject("RequestDict")
					ReqDictTemp.ForceAdd "AsyncAction", "unitselection"
					ReqDictTemp.ForceAdd "PartialRenderControlList", Context.PageControlKey
					ReqDictTemp.ForceAdd "Silent", "1"
					ReqDictTemp.Remove "Debug"
					ReqDictTemp.Remove "PartialRenderType"
					ReqDictTemp.Remove "Trace"
					
					Page.PrintTrace "BookingUI_StayMain_Polling: Render available stays as calendars - start"
					BookingUI_RenderAvailCal pO, objDictAvaiStays, False
					Page.PrintTrace "BookingUI_StayMain_Polling: Render available stays as calendars - end"
					pO.Write "<script type=""text/javascript"">"
					pO.Write "NewMind.ETWP.ControlData[" & Context.PageControlKey & "] = { "
					pO.Write  "UnitSelPartialRenderLink: '"
					pO.Write  Page.Functions.GetNewObject("clsJSON").EscapeJSON("?" & ReqDictTemp.Querystring)
					pO.Write "'"
					pO.Write " };"
					pO.Write "NewMind.ETWP.Booking.InitUnitSel();"
					pO.Write "</"&"script>"
				
				End If 

				' Kick off the show / hide script for fuzzy result sets now that we've rendered out all
				' the content rather than waiting for page load - hopefully we can remove some of the
				' flicker that occurs otherwise
				pO.Write "<script type=""text/javascript"">NewMind.ETWP.Booking.InitPollingUnitSel();</"&"script>"
					
			End Function

			Public Function RenderNotRequiredDateWarning(pO)
				pO.Write "<p class=""fuzzyWarning"">" & Page.Resource("bookonline/unitselection/notrequireddates", "Sorry, we don't have any availability for the dates you requested. These are the nearest available dates for your room and duration requirements.") & "</p>"
			End Function

			Public Function RenderStay(objFuzzyStay, objAvail, intIndex, objRenderSettings, bIsTeleBooking, ByVal strProductBookingWebIfAny, ByVal strEviivoIdIfAny, ByVal objAllUnits)
				
				' 2011-08-09 DWR: Expect the BookingRequirement in objRenderSettings to be read-only (since it usually comes from Page.Functions.GetSharedObject),
				' so replace it with an editable version (since some methods in here try to mess about with properties on it)
				Set objRenderSettings.BookingRequirement = GetEditableBookingRequirement(objRenderSettings.BookingRequirement)

				Dim objSuppliersForStay : Set objSuppliersForStay = objAvail.GetSupplierUnitDataForStay(objFuzzyStay.StartDate, objFuzzyStay.Nights)
			
				Dim intProdKey : intProdKey = objRenderSettings.ProductKey
				Dim dStartNight : dStartNight = objRenderSettings.BookingRequirement.VisitDate
				Dim iNights : iNights = objRenderSettings.BookingRequirement.Nights
				Dim pO : Set pO = objRenderSettings.OutputWriter
				
				'just need to set these here as we may be coming in direct from partial render request
				bRenderAsCalendar = objRenderSettings.RenderAsCalendar
				IsVBPollingEnabled = objRenderSettings.IsVBPollingEnabled
				
				' Loop through each supplier and render their units
				' - Suppliers will be ordered NewMind, FrontDesk, Other
				' - If "Booking_ForceExternal" is enabled, FrontDesk is treated as "Other"
				' - There is a limit on the number of "Other" entries to be rendered (if ForceExternal
				'   is enabled, then FrontDesk counts towards this limit)
				' - If ForceExternal is not enabled, FrontDesk will only be rendered if there is no
				'   local availability
				Dim bStayHasLocalAvail : bStayHasLocalAvail = False
				Dim intExtSuppliersShown : intExtSuppliersShown = 0
				Dim bRenderedStaySummary : bRenderedStaySummary = False
				Dim intIndexSupplier
				Dim objSupplier
				
				Dim bSkipSupplier
				Dim strBookingStaySummary
				Dim bExternalSupplier
				Dim strSupplierId, strSupplierName, strSupplierQuality, strSupplierLogo, strSupplierEviivoName
				Dim intBookingType
				
				Dim bPreciseMatch : bPreciseMatch = (DateValue(objFuzzyStay.StartDate) = DateValue(dStartNight)) And (objFuzzyStay.Nights = iNights)
				
				For intIndexSupplier = 0 To objSuppliersForStay.Count - 1
					
					' Get basic supplier data - count FrontDesk as "Other" if ForceExternal enabled
					Set objSupplier = objSuppliersForStay.GetItem(intIndexSupplier)
					If (objSupplier.IsLocal) Then
						bStayHasLocalAvail = True
					End If
					bExternalSupplier = (objSupplier.IsExternal Or (objSupplier.IsRemote And Page.Site.Params("Booking_ForceExternal")))
					
					' Don't render FrontDesk if got local avail for this stay and not enabled ForceExternal
					bSkipSupplier = (bStayHasLocalAvail And (objSupplier.IsRemote And (Not Page.Site.Params("Booking_ForceExternal"))))
					If Not (bSkipSupplier) Then
						
						' Don't bother rendering stay summary title if we've got a perfect match, as we
						' won't be showing any fuzzy content if there's a spot-on option
						If ((Not bPreciseMatch) And (Not bRenderedStaySummary)) Then
							If Not bRenderAsCalendar Then
								BookingUI_StaySummary dStartNight, iNights, objFuzzyStay.StartDate, objFuzzyStay.Nights, pO
							End If 
							bRenderedStaySummary = True
						End If
						
						' If this is an external supplier, we need the deep-link quality to pass to get
						' included in the hidden booking-info form fields
						
						' PW 2010-07-28 I have added a new field called strSupplierEviivoName
						' This is to pass through the original name field from Eviivo through to the polling exit page.
						' Previously, we did some manipulation on this value to ensure it had a nice display name.
						' However, this had broken Eviivo's own external link - we have in the past asked Eviivo to provide a
						' nice display name field but until they do so we are going to have to do our own and pass both values through as hidden 
						' form fields
						If (bExternalSupplier) Then
							strSupplierId = objSupplier.ID
							strSupplierName = objSupplier.DisplayName
							strSupplierQuality = objSupplier.Quality
							strSupplierEviivoName = objSupplier.Name
						Else
							strSupplierId = Null
							strSupplierName = Null
							strSupplierQuality = Null
							strSupplierEviivoName = Null
						End If
						
						' Render the actual options (wrap in the standard form tag)
						If (objSupplier.IsLocal) Then
							If IsEmpty(IsExternalBooking) Then
								InitExternalBookingSettings()
							End If
							If (IsExternalBooking) Then
								intBookingType = BOOKING_Redirect
								strProductEstateID = DMS.GetProductEstateID(intProdKey)
								strExtBookUrl = GetExtBookUrlFromProductEstate(strProductEstateID)
							Else
								intBookingType = BOOKING_Local
							End If
						ElseIf (objSupplier.IsExternal) Then
							' 2011-07-20 DWR: We don't need to call InitExternalBookingSettings if dealing with an VB Polling product as
							' the next page should always be the Polling Exit (no point redirecting to another site which will then - if
							' it's an NM site - have to display another redirect page to book the product)
							intBookingType = BOOKING_PollingRedirect
						Else
							If IsEmpty(IsExternalBooking) Then
								InitExternalBookingSettings()
							End If
							If (IsExternalBooking) Then
								intBookingType = BOOKING_Redirect
								strProductEstateID = DMS.GetProductEstateID(intProdKey)
								strExtBookUrl = GetExtBookUrlFromProductEstate(strProductEstateID)
							Else
								intBookingType = BOOKING_Eviivo
							End If
						End If
					
						' Local and FrontDesk both use current site name w/out logo
						' External Suppliers should have their own logo passed in
						' PW - 	moved this out of BookingUI_StayDetails_PollingHeader 
						'		we can now pass it to the hidden form fields
						'		for use on the polling exit page
						If (objSupplier.IsExternal) Then
							strSupplierLogo = objSupplier.Logo
							If (Trim("" & strSupplierName) = "") Then
								strSupplierName = "Unnamed Supplier"
							ElseIf (strSupplierLogo = "") Then
								' 2014-07-01 DWR: It's common for Eviivo to not return logo data for the Polling Providers so for most cases we take the Supplier Name (the
								' Eviivo version, rather than the "friendly" version that we maintain) and request the logo from ntop using it. For cases where Eviivo
								' results are treated as Polling results (see FogBugz 10386), we need a special case (the friendly name will always be "Eviivo" in
								' this case).
								If (strSupplierName = "Eviivo") Then
									strSupplierLogo = Page.ImageResource( _
										"bookonline/unitselection/polling/eviivo", _
										"/engine/shared_gfx/eviiopollingresult.jpg" _
									)
								Else
									' 2008-12-09 DWR: Supplier Logo isn't actually going to be received from the Eviivo Component, we mash Supplier Name into this url
									' 2010-03-04 DWR: Eviivo moved the logo location..
									strSupplierLogo = "http://www.ntopsearch.com/media/images/Suppliers/" & strSupplierEviivoName & ".gif"
								End If
							End If
						Else
							' 2009-02-12 DWR: Changed the way in which supplier name and logo are determined for Local / FrontDesk
							' suppliers (ie. the non-external entries) - before it had no logo and displayed the site name, now
							' these are the defaults, but content can be pulled from languages xml. This content can be specified
							' to vary per estate if desired (intended to be used when VB Polling is combined with Force External
							' Bookings)
							' - Supplier Name
							strSupplierName = ""
							If IsExternalBooking Then
								strSupplierName = Page.Resource("bookonline/unitselection/polling/localsupplier/estate_" & strProductEstateID & "/name", "")
							End If
							If (strSupplierName = "") Then
								'#MJ -	the resource manage is the same for both main sites and channel sites
								'		therefore we can never use Page.Site.Name as an alternative value as this would be cached wrongly by the ResourceManager
								'		so try to pull one from there, if not fall back to the site name
								strSupplierName = Page.Resource("bookonline/unitselection/polling/localsupplier/name", "")
								If strSupplierName = "" Then
									strSupplierName = Page.Site.Name
								End If
							End If
							' - Supplier Logo
							strSupplierLogo = GetSupplierLogo(strProductEstateID)
			
						End If
						
						' 2013-02-05 TB: objRenderSettings is used by RenderBookingInfoForm to populate some hidden stay information
						' For fuzzy stays, both nights and startdate may differ from the original requirements.
						' For FogBugz case 7594 I added the second line below which wasn't present.
						objRenderSettings.BookingRequirement.VisitDate = objFuzzyStay.StartDate
						objRenderSettings.BookingRequirement.Nights = objFuzzyStay.Nights

						' 2014-03-12 DWR: We need to pass the Search Industry Classification into the form rendering code for VB Polling Products so that the
						' Polling Exist can generate the deep link correctly. An Eviivo Configset can be set up with zero, meaning support either 1 OR 9. The
						' Avail Component will perform searches for both in that case but only allow any Products to return results for one. Since we won't
						' get an objSupplier reference with zero units (since that would mean it's not got availability and we're only looking at available
						' options here) we can just grab the IndustryClassification values from the first Unit since it is guaranteed to be consistent
						' across all Units for this booking option. The IndustryClassification value will be zero for non-Eviivo data but that won't
						' matter since it's only ever consider in the Polling Exit which is for Eviivo results only.
						RenderBookingInfoForm _
							pO, _
							intProdKey, _
							objRenderSettings, _
							intBookingType, _
							strSupplierId, _
							strSupplierName, _
							strSupplierEviivoName, _
							strSupplierQuality, _
							strSupplierLogo, _
							objSupplier.Units.GetItem(0).IndustryClassification
								
						BookingUI_StayDetails_PollingHeader objSupplier, pO, strSupplierLogo, strSupplierName

						
						' 2009-09-14 DWR: Forcing iStayNum to "1" every time - since we are clearly only having
						' one stay per form (since we open the form above - in RenderBookingInfoForm - and we
						' close it below) we'll always be passing only a single stay to the next stage. This
						' makes things easier - the multiple-stays-per-form idea was ridiculous.
						' 2010-10-21 TB: Changing back to use unique stay index. Multiple stays per form will
						' happen for fuzzy results and calendar view. html ids use the stay key, as does the JS
						' when choosing to show/hide the book now button.
						BookingUI_StayDetails _
							objSupplier, _
							intIndex, _
							dStartNight, _
							iNights, _
							bIsTeleBooking, _
							strProductBookingWebIfAny, _
							strEviivoIdIfAny, _
							objRenderSettings.ProductKey, _
							objRenderSettings.Channel, _
							objFuzzyStay.Indicative, _
							Not objFuzzyStay.HasInvalidIndicative, _
							objAllUnits, _
							Nothing, _
							pO, _
							false 'we can't render the maximum available units for polling

						pO.Write "</form>"

					End If

				Next

			End Function
			
			Private Function RenderNoAvailElement(ByVal objRenderSettings)
				
				Dim pO : Set pO = objRenderSettings.OutputWriter 
				pO.Write "<div class=""pnNoAvail"">"
				pO.Write Page.Resource("bookonline/unitselection/noavailability", "<p>No availability for this product for the specified date. This may occur if the accommodation is booked prior to your arrival at this page.</p>")
				pO.Write "</div>"
				
				If objRenderSettings.RenderAsCalendar Then 
					
					Dim strClassMonth : strClassMonth = "MonthWrapper"
					
					pO.Write "<div class=""CalendarsWrapper"">"
'				
					BookingUI_RenderCalendarMonth pO, objRenderSettings.BookingRequirement.VisitDate, strClassMonth & " currentmonth"
'					' last day + 1 to get the first day of the next month for the calendar
					BookingUI_RenderCalendarMonth pO, Page.Functions.Dates.fn_GetLastDateOfMonth(objRenderSettings.BookingRequirement.VisitDate) + 1, strClassMonth &  " nextmonth"

					' global count used to track how many calendars have been added to the output for the prev/next buttons
					g_iNumberOfCalendarsRendered = 2
					
					BookingUI_RenderAvailCalLinks objRenderSettings.BookingRequirement.VisitDate, pO
					BookingUI_RenderAvailCalKey pO
					pO.Write "</div>"
					
					pO.Write "<script type=""text/javascript"">NewMind.ETWP.Booking.UpdateCalLinks();</" & "script>"
				End If 	
			
			End Function
			
			' ====================================================================================================
			' RENDER: Main entry point when VB Polling is disabled (or handling tickets, not acco products)
			' ====================================================================================================
			Private Function BookingUI_StayMain_Legacy(objData, objRenderSettings)
				' This is the non-VB-Polling approach (supports EITHER FrontDesk OR local availability for accommodation)
				'reset the output variable to our OutputWriter
				Dim pO : Set pO = objRenderSettings.OutputWriter
				
				Dim intBookingType
				Dim bNoResults
				Dim objFuzzyStayOptions, objFuzzyStay
				Dim objSuppliersForStay
				Dim objAvailEntry
				Dim lsRemoteUnitSelections
				
				Dim objAvail: Set objAvail = objData.Availability
				Dim intProdKey: intProdKey = objData.Product_Key
				
				Dim bIsTeleBooking: bIsTeleBooking = objData.IsOnTeleBookingChannel

				' Grab hold of the data (in this method, there should only ever be zero or one fuzzy
				' stay options, as the BookingUI_StayMain_Legacy method handle fuzzy availability)
				Set objFuzzyStayOptions = objAvail.GetUniqueFuzzyCombinations()
				If (objFuzzyStayOptions.Count = 0) Then
					Page.PrintTraceWarning "objAvail.GetUniqueFuzzyCombinations reported zero stay options"
					bNoResults = True
				Else
					' Any suppliers returned here will be sorted with Local / NewMind first, then FrontDesk
					' second (if we have both) - if there are multiple, it should always be the first one
					' that we want
					Set objFuzzyStay = objFuzzyStayOptions.GetItem(0)
					Page.PrintTrace "BookingUI_StayMain_Legacy: Get data for stay - " & objFuzzyStay.StartDate & ", " & objFuzzyStay.Nights
					Set objSuppliersForStay = objAvail.GetSupplierUnitDataForStay(objFuzzyStay.StartDate, objFuzzyStay.Nights)
					If (objSuppliersForStay.Count = 0) Then
						Page.PrintTraceWarning "objAvail.GetSupplierUnitDataForStay reported zero suppliers"
						bNoResults = True
					Else
						Set objAvailEntry = objSuppliersForStay.GetItem(0)
						bNoResults = False
					End If
				End If
				
				' Open form and prepare to wrap content in "staySelection" container
				If (IsExternalBooking) Then
					intBookingType = BOOKING_Redirect
				Else
					intBookingType = BOOKING_Local
				End If
				
				RenderBookingInfoForm pO, intProdKey, objRenderSettings, intBookingType, Null, Null, Null, Null, Null, Null
				
				pO.Write "<div class=""staySelection"">"

				' Render info (or display warning if no availability)
				If bNoResults Then
					RenderNoAvailElement objRenderSettings
					bProdHasAvail = False ' This is exposed through the WSC's public property "ProdHasAvail"
				Else
					bProdHasAvail = True ' This is exposed through the WSC's public property "ProdHasAvail"
					If objRenderSettings.BookingType = "accommodation" Then
				
						' Retrieve any unit selections that have been passed in through the querystring
						' - eg. when VisitBritain hooks in to complete a booking
						' There will be an entry in lsUnitSelections for each requirement.
						' Note that ReqNo in the avail data is one-based while the lsUnitSelections indices
						' are zero-based, so the UnitKey for ReqNo 1 = lsUnitSelections(0). If there was no
						' selection made for a ReqNo, the lsUnitSelections value will be zero.
						' NB: This value might be Nothing if no selections are passed in on querystring.
						Set lsRemoteUnitSelections = BookingUI_UnitSel_GetOptionsRemoteSelected(objAvailEntry)

						' Render the unit selection options (pass "1" as iStayNum parameter - we'll only
						' be rendering a single stay option here, since fuzzy isn't supported in this
						' configuration..)
						BookingUI_StayDetails _
							objAvailEntry, _
							1, _
							objRenderSettings.BookingRequirement.VisitDate, _
							objRenderSettings.BookingRequirement.Nights, _
							bIsTeleBooking, _
							objData.bookingweb, _
							objData.EviivoId, _
							intProdKey, _
							objRenderSettings.Channel, _
							objFuzzyStay.Indicative, _
							Not objFuzzyStay.HasInvalidIndicative, _
							objData.Units, _
							lsRemoteUnitSelections, _
							pO, _
							objRenderSettings.RenderMaximumUnitsAvailable

					Else
						BookingUI_TicketsSummary objAvailEntry, objRenderSettings.BookingRequirement.VisitDate, pO
					End If
				End If
				
				' Close "staySelection" div and form
				pO.Write "</div>"
				pO.Write "</form>"
			End Function

			
			' ==========================================================================================================
			' These functions are all about assigning unit selections when the booking process is hooked into from an
			' external site (eg. VisitBritain).
			'
			' The other site will have requested availability data through the webservice and the order in which
			' the options appear there may vary from the order that they're returned from the availability object's
			' queries.
			'
			' These functions are intended to pick up on selections in the querystring and match them back up to
			' the ReqNo entries.
			'
			' Selections passed in are given as parameters of the form
			'  URslt1=12345,1,1
			' where the comma-separated values are UnitKey, number of adults, number of children.
			' The "1" in "URslt1" is expected to match up with ReqNo 1 when passed through, but the problem is that
			' this often isn't the case.
			' ==========================================================================================================
			
			' SUMMARY: prepare a list of UnitKey selections for each ReqNo is availability recordset
			' [rsAvail]: ADO unit recordset from availability object
			' <retval>: clsList with as many values as there are ReqNo entries, containing the UnitKey for each one
			Private Function BookingUI_UnitSel_GetOptionsRemoteSelected(ByVal objAvailEntry)
				Dim intIndex, objUnit
				Dim arrReqUnitOptions
				Dim arrReqUnitSelections
				Dim intUnitSel
				Dim lsUnitKeys
			
				' Build up a list of unit options:
				' - Will get a list of objects where each object has properties:
				'    > ReqNo (integer)
				'    > NumPeople (integer)
				'    > Units (list of integers)
				' - We're going to loop through the availability recordset, so must remember
				'   to return it back to the beginning when we're done
				Set arrReqUnitOptions = Page.Functions.GetNewObject("clsList")
				For intIndex = 0 To objAvailEntry.Units.Count - 1
					Set objUnit = objAvailEntry.Units.GetItem(intIndex)
					BookingUI_UnitSel_AddReqUnitOption arrReqUnitOptions, objUnit.ReqNo, objUnit.ReqSize, objUnit.UnitKey
					'BookingUI_UnitSel_AddReqUnitOption arrReqUnitOptions, objUnit.ReqNo, objUnit.UnitCount, objUnit.UnitKey
				Next
			
				' Build up a list of unit selections passed in from external site (eg. VisitBritain):
				' - Will get a list of objects where each object has properties:
				'    > NumPeople (integer)
				'    > UnitKey (integer)
				'    > PossReqNos (list of integers)
				'       = list of ReqNo values that this may be
				'         a user selection for
				intUnitSel = 0
				Set arrReqUnitSelections = Page.Functions.GetNewObject("clsList")
				Do
					intUnitSel = intUnitSel + 1
					If (Len(Request("URslt" & intUnitSel)) > 0) Then
						BookingUI_UnitSel_AddReqUnitSelection arrReqUnitSelections, Request("URslt" & intUnitSel), arrReqUnitOptions
					Else
						Exit Do
					End If
				Loop
				
				' If there were no selections passed in like this, return Nothing
				If (arrReqUnitSelections.Count = 0) Then
					Dim BookingUI_UnitSel_GetOptionSelected
					Set BookingUI_UnitSel_GetOptionSelected = Nothing
				End If
				
				' Now try to return matched unit options / selections
				' - Get back a list of unit keys, one key per requirement
				'   (If failed to get a perfect match, some of these values may be zero)
				Set BookingUI_UnitSel_GetOptionsRemoteSelected = BookingUI_UnitSel_GetMatchedReqUnitSelection(arrReqUnitOptions, arrReqUnitSelections)
			
			End Function
			
			Private Function BookingUI_UnitSel_AddReqUnitOption(arrReqUnitOptions, intReqNo, intNumPeople, intUnitKey)
				Dim objEntry, objEntryPrev
			
				' Input list SHOULD be initialised as an empty list, but just in case..
				If IsEmpty(arrReqUnitOptions) Or IsNull(arrReqUnitOptions) Then
					Set arrReqUnitOptions = Page.Functions.GetNewObject("clsList")
				End If
				
				' If we've already got list items, check whether we're still working on the same
				' ReqNo as the previous entry. If so, add to that entry's unit list.
				If (arrReqUnitOptions.Count > 0) Then
					Set objEntryPrev = arrReqUnitOptions(arrReqUnitOptions.Count - 1)
					If (objEntryPrev("ReqNo") = intReqNo) Then
						objEntryPrev("Units").Add intUnitKey
						Exit Function
					End If
				End If
					
				' Need to create a new entry
				Set objEntry = Page.Functions.GetNewObject("clsValueBag")
				objEntry("ReqNo") = intReqNo
				objEntry("NumPeople") = intNumPeople
				Set objEntry("Units") = Page.Functions.GetNewObject("clsList")
				objEntry("Units").Add intUnitKey
				arrReqUnitOptions.Add objEntry
			
			End Function
					
			Private Function BookingUI_UnitSel_AddReqUnitSelection(arrReqUnitSelections, strUnitSelInfo, arrReqUnitOptions)
				Dim arrSegments
				Dim intNumAdults
				Dim intNumChildren
				Dim intUnitKey
				Dim intIndex
				Dim objEntry
				Dim objUnitList

				' Input list SHOULD be initialised as an empty list, but just in case..
				If IsEmpty(arrReqUnitSelections) Or IsNull(arrReqUnitSelections) Then
					Set arrReqUnitSelections = Page.Functions.GetNewObject("clsList")
				End If
				
				' strUnitSelInfo should be of the form "UnitKey,NumAdults,NumChildren"
				' Exit if not
				arrSegments = Split(strUnitSelInfo, ",")
				If (UBound(arrSegments) <> 2) Then
					Exit Function
				End If
				
				' Ensure entries in string are numeric (exit if not)
				On Error Resume Next
				intUnitKey = CLng(arrSegments(0))
				If (Err) Then
					Exit Function
				End If
				intNumAdults = CLng(arrSegments(1))
				If (Err) Then
					Exit Function
				End If
				intNumChildren = CLng(arrSegments(2))
				If (Err) Then
					Exit Function
				End If
				On Error Goto 0
			
				' Ensure values look reasonable
				If ((intUnitKey <= 0) _
				 Or (intNumAdults < 0) Or (intNumChildren < 0) _
				 Or ((intNumAdults + intNumChildren) <= 0)) Then
			 		Exit Function
				End If
			
				' Preparer new entry
				Set objEntry = Page.Functions.GetNewObject("clsValueBag")
				objEntry("NumPeople") = intNumAdults + intNumChildren
				objEntry("UnitKey") = intUnitKey
				Set objEntry("PossReqNos") = Page.Functions.GetNewObject("clsList")
				
				' Look through the unit options and look for possible requirement matches
				' - We've got a set of requirement / room options from the DMS and we've (possibly) got a
				'   set of unit selections from VisitBritain (or whoever), but these may not currently be
				'   aligned, so we want to determine the possible ways they MIGHT go together, and we'll
				'   try to get the best configuration (which will hopefully match the original choice)
				'   later on.
				If (arrReqUnitOptions.Count > 0) Then
					For intIndex = 0 To arrReqUnitOptions.Count - 1
						' If requirement option matches the selection's NumPeople and contains the
						' UnitKey, then we've got a possible match
						If ((arrReqUnitOptions(intIndex)("NumPeople") = objEntry("NumPeople")) _
						And (arrReqUnitOptions(intIndex)("Units").Contains(objEntry("UnitKey")))) Then
							objEntry("PossReqNos").Add arrReqUnitOptions(intIndex)("ReqNo")
						End If
					Next
				End If
				
				' If there is at least one possible requirement match, add entry to list
				' (Otherwise, we can't do anything with the selection so don't bother with it)
				If (objEntry("PossReqNos").Count > 0) Then
					arrReqUnitSelections.Add objEntry
				End If
				
			End Function
		
			Private Function BookingUI_UnitSel_GetMatchedReqUnitSelection(arrReqUnitOptions, arrReqUnitSelections)
				' Given list of requirement option objects and unit selection objects, try to match them up.
			
				Dim lsPermutations
				Dim lsTemp
				Dim lsPossReqNos
				Dim intIndex
				Dim intIndexSel, intIndexPoss, intIndexPerm
				Dim intIndexOption
			
				Dim intScore, intBestScore, strBestPermutation
				
				Dim arrMatches
				Dim intUnitKey
				Dim lsUnitKeys
				
				Dim GetMatchedReqUnitSelection
				
				' Ensure we've got values for both lists
				If IsNull(arrReqUnitOptions) Or IsNull(arrReqUnitSelections) Then
					Set GetMatchedReqUnitSelection = Nothing
				End If
				If (arrReqUnitOptions.Count = 0) Or (arrReqUnitSelections.Count = 0) Then
					Set GetMatchedReqUnitSelection = Nothing
				End If
			
				' First, create a list of ways in which the unit selections could be applied to the unit
				' options. We'll get out a list of strings which are comma-separated lists; the values
				' will relate the arrReqUnitSelections list indices to arrReqUnitOptions entries.
				'  eg. string "2,3,1"
				'      maps Selection 1 -> Option 2
				'           Selection 2 -> Option 3
				'           Selection 3 -> Option 1
				Set lsPermutations = Page.Functions.GetNewObject("clsList")
				For intIndexSel = 0 To (arrReqUnitSelections.Count - 1)
					Set lsPossReqNos = arrReqUnitSelections(intIndexSel)("PossReqNos")
					If (lsPermutations.Count = 0) Then
						' This is the first pass, so initialise the permutations list with
						' the possible matches from this first ReqUnitSelection
						For intIndexPoss = 0 To (lsPossReqNos.Count - 1)
							lsPermutations.Add lsPossReqNos(intIndexPoss)
						Next
					Else
						' We want to take our whatever permutation strings we have so far and expand
						' them to include the possibilities for this ReqUnitSelection
						' - Make a copy of lsPermutations thus far
						Set lsTemp = Page.Functions.GetNewObject("clsList")
						For intIndexPerm = 0 To (lsPermutations.Count - 1)
							lsTemp.Add lsPermutations(intIndexPerm)
						Next
						' - Clear out permutation list
						lsPermutations.Clear
						' - Re-create new list using previous values with new combinations
						For intIndexPoss = 0 To (lsPossReqNos.Count - 1)
							For intIndexPerm = 0 To (lsTemp.Count - 1)
		    				lsPermutations.Add lsTemp(intIndexPerm) & "," & lsPossReqNos(intIndexPoss)
		    			Next
		    		Next
					End If
				Next
			
				' Now determine which arrangement matches the most selection / options pairs
				intBestScore = -1
				For intIndex = 0 To (lsPermutations.Count - 1)
					intScore = BookingUI_UnitSel_ScoreUnitSelPermutation(lsPermutations(intIndex))
					If (intScore > intBestScore) Then
						intBestScore = intScore
						strBestPermutation = lsPermutations(intIndex)
					End If
				Next
				
				' Finally, translate these matches into UnitKey values (or zero for unit
				' option which don't have a selection matched to them)
				' - Start off with a full-size list (matching size of arrReqUnitOptions) with
				'   with all zero values
				Set lsUnitKeys = Page.Functions.GetNewObject("clsList")
				For intIndex = 0 To (arrReqUnitOptions.Count - 1)
					lsUnitKeys.Add 0
				Next
				
				' - Now push in the selection matches we have
				'    > Split best permutation back into integer values in arrMatches
				'    > The index of arrMatches will matches the index of arrReqUnitSelections
				'    > The value of arrMatches(n) will be the ReqNo it matches, which is the index
				'      of arrReqUnitOptions + 1 (andso also the index of lsUnitKeys + 1 since these
				'      two lists overlay)
				arrMatches = Split(strBestPermutation, ",")
				For intIndexSel = 0 To UBound(arrMatches)
					intIndexOption = arrMatches(intIndexSel) - 1
					intUnitKey = arrReqUnitSelections(intIndexSel)("UnitKey")
					lsUnitKeys(intIndexOption) = intUnitKey
				Next
				
				' Return matches!
				' There are the same number of values in lsUnitKeys as in arrReqUnitSelections, and
				' each lsUnitKeys(n) is the UnitKey for arrReqUnitSelections(n)
				Set BookingUI_UnitSel_GetMatchedReqUnitSelection = lsUnitKeys
				
			End Function	
			
			Private Function BookingUI_UnitSel_ScoreUnitSelPermutation(strPermutation)
				' Determine a score for the Unit Selection / Option permutations calculated above.
				' Basically, give a score of one for each non-duplicated match.
			
				Dim intIndex
				Dim intScore
				Dim arrValues
				Dim lsReqNos
			
				Set lsReqNos = Page.Functions.GetNewObject("clsList")
				arrValues = Split(strPermutation, ",")
				intScore = 0
				For intIndex = 0 To UBound(arrValues)
					If Not lsReqNos.Contains(arrValues(intIndex)) Then
						intScore = intScore + 1
						lsReqNos.Add arrValues(intIndex)
					End If
				Next
			
				BookingUI_UnitSel_ScoreUnitSelPermutation = intScore
			
			End Function
					

			' ====================================================================================================
			' RENDER: Render options for accommodation products (only used with non-precise fuzzy stays)
			' ====================================================================================================
			' SUMMARY: summarise STAYS for this product which match booking criteria
			' [arsAvail]: ADO unit recordset from availability object
			' [adtStartNight]: date of first night of stay
			' [aiReqNumNights]: integer requested num nights
			Private Function BookingUI_StaySummary( _
				dtReqFirstNight, _
				iReqNights, _
				dtStayFirstNight, _
				iStayNights, _
				pO)
				
				' Render each stay result with link to further details
				' - 2009-08-10 DWR: Why do we not render this if "_stay" is in the querystring???
				If (Request("_stay") <> "") Then
					Exit Function
				End If

				pO.Write "<div class=""StayCandidateList"">"
				pO.Write  "<div class=""StayCandidatesTtl"">"
				pO.Write   "<p>" & Page.Resource("bookonline/unitselection/flexiblesearchresults", "Flexible Search Results") & "</p>"
				pO.Write  "</div>"
				If (dtStayFirstNight <> dtReqFirstNight) Or (iReqnights <> iStayNights) Then
					pO.Write  "<div class=""cell"">"
					pO.Write "<div class=""pnStayTtl"">"
					pO.Write BookingUI_StayTtl(dtStayFirstNight, iStayNights)
					pO.Write "</div>"
					pO.Write BookingUI_StayDiff(dtReqFirstNight, dtStayFirstNight, iReqNights, iStayNights)
					pO.Write  "</div>"
				End If
				pO.Write "</div>"

			End Function

			' SUMMARY: Render details for a single stay, including UNIT booking UI
			' [objAvailEntry]: A single supplier's availability data for a single stay (AvailabilityStayResultsWrapped)
			' [iStayNum]: Only applies when displaying multiple fuzzy results
			' [adtStartNight]: date of first night of stay
			' [aiReqNights]: integer requested num nights
			' [bTeleBooking]: does the current product only support telephone booking (ie. is on tele booking channel)?
			' [strProductBookingWebIfAny]: the Booking Website for the the current product, if there is one (so may be empty, null, blank, whatever)
			' [strEviivoIdIfAny]: the Eviivo Id for the the current product, if there is one (so may be empty, null, blank, whatever)
			' [intProductKey]
			' [strChannel]
			' [bIndicative]: does the specified stay have any indicative units?
			' [bIndicativeValid]: are we within the timeout period for indicative bookings?
			' [lsRemoteUnitSelections]: data regarding unit pre-selections (see VB Deep Linking)
			Private Function BookingUI_StayDetails(_
				ByVal objAvailEntry, _
				ByVal iStayNum, _
				ByVal adtStartNight, _
				ByVal aiReqNights, _
				ByVal bTeleBooking, _
				ByVal strProductBookingWebIfAny, _
				ByVal strEviivoIdIfAny, _
				ByVal intProductKey, _
				ByVal strChannel, _
				ByVal bIndicative, _
				ByVal bIndicativeValid, _
				ByVal objAllUnits, _
				ByVal lsRemoteUnitSelections, _
				ByVal pO,_
				ByVal bRenderMaximumUnitsAvailable)
				
				Dim intIndexUnit, objUnit
				Dim iLastReqmnt, iThisReqmnt
				Dim bGotOpenReqContainer
				Dim sClassName, bPrecise, iUnitKey
				Dim iMaxRq, iRemoteUnitKey
				Dim bSelected
				Dim strNonBookableUnits, bHasBookableUnits, bHasNonBookableUnits
				
				' Ensure we've actually got some availability (we should if we've got here!)
				If (objAvailEntry.Units.Count = 0) Then
					Page.PrintTraceWarning "BookingUI_StayDetails: No units in objAvailEntry"
					Exit Function
				End If
			
				' This method opens a new div - we'll need to close it later
				BookingUI_RenderNewStay objAvailEntry, iStayNum, adtStartNight, aiReqNights, pO
				
				iMaxRq = 0
				iLastReqmnt = 0
				iRemoteUnitKey = 0
				bGotOpenReqContainer = False
				bHasBookableUnits = False
				bHasNonBookableUnits = False
				For intIndexUnit = 0 To objAvailEntry.Units.Count - 1
					Set objUnit = objAvailEntry.Units.GetItem(intIndexUnit)
					
					iThisReqmnt = objUnit.ReqNo
					If iThisReqmnt > iMaxRq Then
						' Moved on to next requirement, get key of pre-selected unit - iRemoteUnitKey
						' will be zero if no selection has been passed in (applies to deep-linking)
						iMaxRq = iThisReqmnt
						iRemoteUnitKey = BookingUI_GetPreSelectedUnitKey(lsRemoteUnitSelections, iThisReqmnt)
					End If
					
					' Check whether we're moving into a new requirement (if so, default to having
					' the first unit appear selected) and render the "Room 1 - for 1 Guest"
					' content
					If iThisReqmnt <> iLastReqmnt Then
						
						' If we've already got one of these containers open, close its tags
						If (bGotOpenReqContainer) Then
							pO.Write "</div></div>"
						End If
						BookingUI_RenderNewReq objUnit, iStayNum, iThisReqmnt, (Not objAvailEntry.IsLocal), pO
						bGotOpenReqContainer = True
						
						bSelected = True
						iLastReqmnt = iThisReqmnt
					Else
						bSelected = False
					End If

					' .. however, if there was a pre-selected unit key passed in, this should override which
					' unit appears selected (this only applies when iRemoteUnitKey is not zero, meaning that
					' a unit selection exists - note: eviivo units always appear with unit key zero)
					iUnitKey = objUnit.UnitKey
					If (iRemoteUnitKey <> 0) Then
						bSelected = (iUnitKey = iRemoteUnitKey)
					End If
					
					' build up a list of invalid indicative or telephone booking
					' units, this is used later by javascript when we have a mixture of allocated and indicative 
					' availability
					If ((objUnit.Indicative And Not bIndicativeValid) Or bTeleBooking) Then
						bHasNonBookableUnits = True
						If Len(strNonBookableUnits) > 0 Then 
							strNonBookableUnits = strNonBookableUnits & ","
						End If
						
						'MJ - 	the stay num is no longer part of this data, it is part of each array's name
						'		look at TB's other changes to see the reasoning behind this
						strNonBookableUnits = strNonBookableUnits & iUnitKey
						Page.PrintTrace "strNonBookableUnits" & strNonBookableUnits
					Else
						bHasBookableUnits = True
					End If

					' 2009-09-30 DWR: The AvailClassName was previously generated by considering the indicative
					' state of the whole stay - this was causing all units to be rendered as indicative if any
					' one of them was, now we take the indicative state from each unit (but keep the indicative
					' "validity" from the whole stay, where required)
					BookingUI_RenderUnit _
						iStayNum, iThisReqmnt, _
						bSelected, _
						objAvailEntry, _
						objUnit, _
						objAllUnits, _
						BookingUI_AvailClassName(objUnit.Indicative, bIndicativeValid, bTeleBooking), _
						pO, _
						bRenderMaximumUnitsAvailable
						
				Next
				
				' Ensure any open req container (eg. "Room 1 - for 1 Guest" section) is closed
				If (bGotOpenReqContainer) Then
					pO.Write "</div></div>"
					bGotOpenReqContainer = False
				End If
				
				' Close the BookingUI_RenderNewStay containing div
				pO.Write "</div>"

				' Wrap these hidden inputs in a div for html validity
				pO.Write "<div>"
				pO.Write "<input type=""hidden"" name=""_nStays"" value=""" & iStayNum & """ />"
				pO.Write "<input type=""hidden"" name=""_nReqs"" value=""" & iMaxRq & """ />"
				If Not (objAvailEntry.IsLocal) Then
					pO.Write "<input type=""hidden"" name=""IsEviivoBooking"" value=""yes"" />"
					If IsExternalBooking Then
						pO.Write "<input type=""hidden"" name=""eviivoconf"" value=""" & CLng("0" & Page.Site.Params("Integration_Eviivo_ConfigSet")) & """ />"
					End If
				End If
				pO.Write "</div>"
				
				' 2014-06-25 DWR: For sites that use the legacy "eviivo external" booking integration (meaning sites where VB Polling is not enabled - the new implementation
				' results in Eviivo results being reported as Polling results and the user being sent through the Polling Exit with a fully-populated deep link), the Book
				' button should not be shown here. The Unit Selection should never be shown in this case, to be honest, since Book buttons should go straight to the Product's
				' Booking Website and not enter the site's availability process. However, if there are sites that show inline Unit Selection (inline with the Product List)
				' then the Unit data may be useful. If we were wanted to render Book buttons here (to the external site) then logic would have to be duplicated from the
				' Product List or Detail Control, which would be better avoided. A much better solution is to enable VB Polling and avoid this legacy mechanism entirely.
				' Note: We could potentially render the button for Local Avail and not for Eviivo but I think that that's more confusing than helpful, particularly since
				' it's inconsistent with the Product List / Detail implementation (which bases its decision upon whether the Product has an Eviivo Id).
				If (Not Page.Site.Params("Booking_EnablePolling")) And (Trim("" & strEviivoIdIfAny) <> "") And Page.Site.Params("Integration_Eviivo_ExtBooking_Enable") Then
					Page.PrintTraceWarning "Not rendering any Book buttons for Unit Selection since the legacy Eviivo External Booking configuration is enabled (the recommended alternative is to use the deep-link-supporting Eviivo External Booking configuration, this may be done by enabling VB Polling)"
					Exit Function
				End If
				
				' 2014-03-14 DWR: New functionality "Availability Searches with offsite Booking Web Booking" allows for Products to be on the Telephone Booking Channel
				' and have their availability queried but to show a Booking button that goes to the Product's Booking Website (if one is specified), rather than
				' showing a "this can not be booked online, please call.." message (this means that the avail criteria have to be re-entered on the target
				' website, but that is understood and how it works - see FogBugz 10367). I've tried to make the markup for this button reminiscent of
				' that in Product List and Detail to try to make any additional styling requirements as low as possible.
				strProductBookingWebIfAny = Trim("" & strProductBookingWebIfAny)
				If (bTeleBooking And Page.Site.Params("Booking_EnableByPhone") And Page.Site.Params("Booking_AllowOffSiteTelephoneBookings") And (strProductBookingWebIfAny <> "")) Then
					Page.PrintTrace "Since this is a Telephone Booking Product with a Booking Website and the 'Allow Offsite Booking Web Booking for Telephone Bookings' parameter is enabled, a button to the Booking Website is being rendered"
					pO.Write "<div class=""pnStayButtons"">"
					pO.Write "<p class=""bookonline"">"
					pO.Write "<a href="""
					pO.Write Server.HtmlEncode(strProductBookingWebIfAny)
					pO.Write """"
					If Page.IsPartialRender Or (Request("PartialRenderType") = "html") Then
						' If in Partial Render then set target="_blank" instead of rel="external" (we only do the latter for strict adherence to standards and then
						' use javascript to transform after rendering - when requesting additional content through javascript this transformation won't be performed
						' so we'll need to generate it direct)
						' 2014-06-12 DWR: The partial render requests for this data are commonly made as "html" meaning that Page.IsPartialRender will be false
						' (the logic being that Controls should render entirely as standard when in html partial render mode) so I've added an additional check
						' for the a "PartialRenderType" value of "html" to ensure that the new-window logic is maintained correctly.
						pO.Write " target=""_blank"""
					Else
						pO.Write " rel=""external"""
					End If
					pO.Write " class=""ProvClickCustom"" name=""PROBWEBREF|" ' This is the "Provider Booking Website Referral" statistic, as required by the SharePoint document for FogBugz 10367
					pO.Write Server.HtmlEncode(strChannel)
					pO.Write "|"
					pO.Write intProductKey
					pO.Write """"
					pO.Write ">"
					pO.Write "<img src="""
					pO.Write Page.ImageResource("bookonline/btn/book", Context.ImageDir & "booking/book.gif")
					pO.Write """ alt="""
					pO.Write Page.Resource("bookonline/btn/book","Book")
					pO.Write " ("
					pO.Write Page.Resource("productdetail/bookonline/opensinanewwindow", "opens in a new window")
					pO.Write ")"" "
					pO.Write "/>"
					pO.Write "</a>"
					pO.Write "</p>"
					pO.Write "</div>" & vbCrLf
					Exit Function
				End If

				' 2014-03-13 DWR: If there is at least one bookable unit then display the Book button and rely on JavaScript to show/hide it if selections are made that
				' can not be completed online. But if there are NO bookable units (eg. a Telephone Booking Product or all of the Units are Indicative where the timeout
				' period has passed) then there's no point even rendering the button.
				If (bHasBookableUnits) Then
					BookingUI_RenderButtons iStayNum, pO, objAvailEntry.IsExternal
				End If
				
				' if we have an invalid indicative unit or telephone unit then
				' render this message - let the js do the rest
				If bHasNonBookableUnits Then
					
					' 2010-07-09 PW: RIP Gary
					' This is the array formerly known as garyTeleBookUnitKeys
					' it is used for switching between the online book button if the unit is bookable
					' or rendering the relevant warning message if it isn't
					' 2010-10-21 TB: augmenting Gary with stay key. This is to allow for multiple stays
					' in which this JS is executed on a per stay basis via a partial render.
					pO.Write "<script type=""text/javascript"">" & vbCrLf
					pO.Write " var aryNonBookableUnits_" & iStayNum & " = [" & strNonBookableUnits & "]; " & vbCrLf
					pO.Write " var iTotalNonBookableUnits = " & iThisReqmnt & ";" & vbCrLf
					pO.Write "</" & "script>"
				
					' Render relevant offline booking message
					pO.Write "<div id=""pnTeleBook_PromptCall"">"
					If Page.Site.Params("Booking_EnableByPhone") Then
						pO.Write "<p>"&Replace(Page.Resource("bookonline/unitselection/telebook/prompt", "One or more of the units you have selected must be booked via telephone. Please ring #bookingtelephone# to continue this booking."),"#bookingtelephone#",Page.Site.Params("Booking_TelephoneNumber"))&"</p>"
					Else
						pO.Write "<p>"&Replace(Page.Resource("bookonline/unitselection/indtelebook/prompt", "Although available, some of the units you have selected cannot be booked online. Alternatively, select different units with online booking only."),"#bookingtelephone#",Page.Site.Params("Booking_TelephoneNumber"))&"</p>"
					End If
					pO.Write "</div>"
				End If

			End Function
			

			Private Function BookingUI_GetPreSelectedUnitKey(ByVal lsRemoteUnitSelections, ByVal iReqNo)
				' remote referrals (eg. VB integration) will include a UNIT_KEY choice and CHILD_COUNT.
				' eg. Request vars formatted as 'URslt[REQ_NUMBER]=[UNIT_KEY]-[NUM_ADULT]-[NUM-CHILD]'

				' If not got here from a remote referral (eg. VB integration), the lsRemoveUnitSelections will be Nothing
				If Not (lsRemoteUnitSelections is Nothing) Then
					' Get unit selection passed in (may be zero if invalid request was made)
					' Note: lsRemoteUnitSelections has zero-based index, iReqNo is one-based
					If ((iReqNo >= 1) And (iReqNo <= lsRemoteUnitSelections.Count)) Then
						BookingUI_GetPreSelectedUnitKey = lsRemoteUnitSelections(iReqNo - 1)
						Exit Function
					End If
				End If

				BookingUI_GetPreSelectedUnitKey = 0
			End Function

			' SUMMARY: for VB Polling - we want to render a supplier name and icon above each set of unit options
			Private Function BookingUI_StayDetails_PollingHeader(ByVal objAvailEntry, ByVal pO, ByVal strSupplierLogo, ByVal strSupplierName)

				' Render header content (icon, if specified) and supplier name
				' 2008-12-18 DWR: Add a style to indicate whether supplier is Local, FrontDesk or External (this will
				' allow a custom logo to be used for Local or FrontDesk, for example)
				pO.Write "<div class=""StayCandidateItemHeader "
				If (objAvailEntry.IsLocal) Then
					pO.Write " AvailLocal"
				ElseIf (objAvailEntry.IsRemote) Then
					pO.Write " AvailFrontDesk"
				Else
					pO.Write " AvailExternal"
				End If
				pO.Write """>"
				If (strSupplierLogo <> "") Then
					pO.Write "<img src=""" & strSupplierLogo & """ alt=""" & strSupplierName & """ />"
				End If
				pO.Write "<h2>" & strSupplierName & "</h2>"
				pO.Write "</div>" & vbCrLf
			End Function
			
			'tries to get a supplier logo for us
			Private Function GetSupplierLogo(strProductEstateID)
				Dim strSupplierLogo : strSupplierLogo = ""
				If IsExternalBooking Then
					strSupplierLogo = Page.ImageResource("bookonline/unitselection/polling/localsupplier/estate_" & strProductEstateID & "/logo", "")
					If (strSupplierLogo = "") Then
						strSupplierLogo = Page.Resource("bookonline/unitselection/polling/localsupplier/estate_" & strProductEstateID & "/logo", "")
						If strSupplierLogo <> "" Then
							Page.PrintTraceWarning "Loaded estate scoped supplier logo from a deprecated location - please move it to the image resources language file"
						End If
					End If
				End If
				If (strSupplierLogo = "") Then
					strSupplierLogo = Page.ImageResource("bookonline/unitselection/polling/localsupplier/logo", "")
					If (strSupplierLogo = "") Then
						strSupplierLogo = Page.Resource("bookonline/unitselection/polling/localsupplier/logo", "")
						If strSupplierLogo <> "" Then
							Page.PrintTraceWarning "Loaded estate scoped supplier logo from a deprecated location - please move it to the image resources language file"
						End If
					End If
				End If
				GetSupplierLogo = strSupplierLogo
			End Function

			' SUMMARY: return URL which browsers without Javascript can use to navigate stay candidates page
			' [aiStay]: integer stay number. 1 = 1st stay, 2 = 2nd stay. Zero produces back URL to stay candidates page
			' <retval>: string URL for hyperlink
			Private Function BookingUI_StayDetailsUrl(aiStay)
				Dim sUrl,sStay,iPos,sRight

				' get current URL. Prepare [stay] variable to be appended to URL
				sUrl=Request.ServerVariables("HTTP_X_REWRITE_URL")
				If aiStay>0 Then
					sStay="&_stay="&aiStay
				Else
					sStay=""
				End If

				' does URL already have stay variable? if so, remove it and return new URL
				iPos=Instr(sUrl,"&_stay=")
				If iPos>0 Then
					sRight=Mid(sUrl,iPos+7)
					iRight=Instr(sRight,"&")
					sUrl=Left(sUrl,iPos-1)
					If iRight>0 Then sUrl=sUrl&Mid(sRight,iRight)
				End If
				BookingUI_StayDetailsUrl=sUrl&sStay
			End Function

			' SUMMARY: render new stay UI - WARNING: this doesn't close all of the elements it opens!
			' [objAvailEntry]: avail data for a single stay
			' [aiStayNum]: integer stay index (1-based)
			' [adtStartNight]: date requested start night
			' [aiReqNights]: integer requested num nights
			Private Function BookingUI_RenderNewStay(ByVal objAvailEntry, ByVal aiStayNum, ByVal adtStartNight, ByVal aiReqNights, ByVal pO)
				Dim sPostfix, bPrecise
				
				' Render slightly differently if got a precise match
				' - Also render differently when VB Polling enabled, since we have to render
				'   more of these sections than otherwise
				Dim bExactMatch : bExactMatch = ((objAvailEntry.StartDate = adtStartNight) And (objAvailEntry.Nights = aiReqNights))
				If (bExactMatch) Or (IsVBPollingEnabled) Then
					bPrecise = True
					sPostfix = "1"
				ElseIf CLng("0" & Request("_stay")) = aiStayNum Then
					sPostfix = "1"
				Else
					sPostfix = ""
				End If
				
				' If not exact match then render a warning as well as the date difference later
				If Not bExactMatch Then
					RenderNotRequiredDateWarning pO
				End If

				pO.Write "<div class=""StayCandidateItem" & sPostfix & """>" & vbCrLf

				If Not bExactMatch Then
					pO.Write "<div class=""pnStayTtl"">"
					pO.Write "<p>"
					pO.Write BookingUI_StayTtl(objAvailEntry.StartDate, objAvailEntry.Nights)
					pO.Write "</p>"
					pO.Write "</div>" & vbCrLf
					If Not bRenderAsCalendar Then 
						pO.Write BookingUI_StayDiff(adtStartNight, objAvailEntry.StartDate, aiReqNights, objAvailEntry.Nights)
					End If 
				End If
			End Function

			' SUMMARY: return title for this stay candidate
			' [aiNights]: integer number nights for this stay
			' [adtFirstNight]: date of first night
			' [adtLastNight]: date of last night
			' <retval>: string stay title
			Private Function BookingUI_StayTtl(ByVal adtFirstNight, ByVal aiNights)
				If aiNights = 1 Then
					BookingUI_StayTtl = _
						aiNights & Page.Resource("bookonline/unitselection/nightstart", " night, start ") & _
						Page.Functions.Dates.ShortDate(adtFirstNight)
					Exit Function
				End If

				BookingUI_StayTtl = _
					aiNights & _
					Page.Resource("bookonline/unitselection/nightsfrom", " nights, from ") & _
					Page.Functions.Dates.ShortDate(adtFirstNight) & _
					Page.Resource("bookonline/unitselection/to", " to ") & _
					Page.Functions.Dates.Shortdate(DateAdd("d", aiNights, adtFirstNight))
			End Function

			' SUMMARY: describe difference between THIS DATE and REQUESTED stay date
			' [adtReqDate]: date of REQUESTED first night of stay
			' [adtThisDate]: date of RESULTANT first night of stay
			' [aiReqNights]: integer requested num nights
			' [aiNights]: integer result num nights
			Private Function BookingUI_StayDiff(ByVal adtReqDate, ByVal adtThisDate, ByVal aiReqNights, ByVal aiResultNights)
				Dim iDateDiff, iDurDiff

				iDateDiff = DateDiff("d", adtReqDate, adtThisDate)
				iDurDiff = aiResultNights - aiReqNights
				BookingUI_StayDiff = _
					"<div class=""pnStayDiff"">" & _
						 Page.Functions.Booking.Booking_MatchQual(0, iDateDiff, iDurDiff, aiReqNights, 2) & _
					"</div>" & vbCrLf
			End Function

			' SUMMARY: render new requirement UI - WARNING: this doesn't close all of the elements it opens!
			' [arsAvail]: ADO unit recordset from availability object
			' [aiStayNum]: integer stay index
			' [aiThisReqmnt]: integer requirement number (from recordset)
			Private Function BookingUI_RenderNewReq(ByVal objUnit, ByVal aiStayNum, ByVal aiThisReqmnt, ByVal abRemote, ByVal pO)

				Dim iSz, sReqmntSet, sRemoteRqmnt
				Dim iChild, iRemoteNumChild, sSelected
				
				
				
				iSz = objUnit.ReqSize

				pO.Write "<div class=""pnStayReqmnt"">" & vbCrLf
				pO.Write "<div class=""pnStayReqmntTtl"">"
				pO.Write Page.Resource("bookonline/unitselection/room","Room")
				pO.Write " "
				pO.Write aiThisReqmnt
				pO.Write " - "
				pO.Write Page.Resource("bookonline/unitselection/for","for")
				pO.Write " "
				pO.Write iSz
				pO.Write " "
				pO.Write Page.Resource("bookonline/unitselection/guest(s)","guest(s)")
				
				'#MJ -	We can only render our room requirement data based upon the recieved dat, not the requirement we passed in, as it may have been fulfilled in a different order
				'2012-03-29 NP: Here we render the requirements that are linked to the unit stay details in the response from the Avail Component
				' we do NOT want to render the original request against each unit that is rendered because they may not order up
				' Example: Request roomReq_1 = 2; roomReq_2 = 1; Response may come back in a different order 
				' i.e. unit_1 with ReqSize = 1, unit_2 with ReqSize = 2 so roomReq_1 = 1, roomReq= 2; they end up swapped around
				pO.Write "<input type=""hidden"" name=""roomReq_" & aiThisReqmnt & """ value=""" & iSz & """ />"
				
				'#MJ - need to check with Rich if we want to indicate who's going into what room
				If Page.Site.Params("Booking_ChildPricing") And objUnit.ChildrenRequirement > 0 Then
					pO.Write " - ("
					pO.Write "<span class=""ReqmntDetails"">"
					pO.Write Page.Resource("adults","Adults")
					pO.Write ": "
					pO.Write objUnit.AdultsRequirement
					pO.Write " "
					pO.Write Page.Resource("children","Children")
					pO.Write ": "
					pO.Write objUnit.ChildrenRequirement
					pO.Write ") "
					pO.Write "</span>"
					' NP 2012-03-01: Child pricing requirements were not previously being posted to the checkout
					' Adult & Child Requirement amount is needed by the RequirementSummary control and the child ages are
					' needed by the checkout for creating the correct requirement record with the relevant discount values
					pO.Write "<input type=""hidden"" name=""roomReq_"& aiThisReqmnt & "_adults"" value=""" & objUnit.AdultsRequirement & """ />"
					pO.Write "<input type=""hidden"" name=""roomReq_"& aiThisReqmnt & "_children"" value=""" & objUnit.ChildrenRequirement & """ />"

					' ChildrenAges is a comma separated list of ages or "", Split will give an empty array if this property is ever Empty
					Dim aryChildAges : aryChildAges = Split(objUnit.ChildrenAges,",")
					Dim iChildAgeIndex : For iChildAgeIndex = 0 To UBound(aryChildAges)
						pO.Write "<input type=""hidden"" name=""roomReq_"& aiThisReqmnt & "_children_childage" & iChildAgeIndex & """ value=""" & aryChildAges(iChildAgeIndex) & """ />"
					Next
					
				End If
				
				pO.Write "</div>" & vbCrLf
				pO.Write "<div class=""pnStayReqmntRslts"">" & vbCrLf
				
			End Function

			' SUMMARY: render unit option HTML
			' [aiStayNum]: integer stay index
			' [aiThisReqmnt]: integer requirement index
			' [aiUnitKey]: integer unit key
			' [bSelected]: should the current unit appear selected 
			' [arsAvail]: ADO availability recordset
			' [asAvailClassName]: string avail class name
			Private Function BookingUI_RenderUnit( _
				ByVal aiStayNum, _
				ByVal aiThisReqmnt, _
				ByVal bSelected, _
				ByVal objAvailEntry, _
				ByVal objUnit, _
				ByVal objAllUnits, _
				ByVal asAvailClassName, _
				ByVal pO, _
				ByVal bRenderMaximumUnitsAvailable)
				
				Dim mUnitStayTotal, iNumNights, mUnitPerNight, iNumPeople, iDaysBreakfast, bPerPerson, mPersonPerNight, strIptId
				Dim mUnitStayTotalPayableBasedOnGuidePrice , bDiscountApplied
				Dim iAdults, iChildren
				
				
				mUnitStayTotal = objUnit.StayTotalPayable
				mUnitStayTotalPayableBasedOnGuidePrice = objUnit.StayTotalPayableBasedOnGuidePrice
				iNumNights = objAvailEntry.Nights
				mUnitPerNight = mUnitStayTotal / iNumNights
				bPerPerson = objUnit.Perperson
				iNumPeople = objUnit.ReqSize
			
				iDaysBreakfast = objUnit.DaysBreakfast
				bDiscountApplied = objUnit.IncludesChildDiscount
				
				Dim iMaxUnitsAvailable : iMaxUnitsAvailable = objUnit.MaximumQuantityAvailable
				
				' We need an id so we can set the label's "for" attribute, but if VB Polling is enabled,
				' we might end up with id duplication - so in that case we append a random suffix
				strIptId = "unit_" & aiStayNum & "_" & aiThisReqmnt & "_" & objUnit.UnitKey
				If (IsVBPollingEnabled) Then
					strIptId = strIptId & "_" & Int(Rnd * 100000)
				End If

				pO.Write "<div class=""pnUnitOption"">"
				pO.Write "<input type=""radio"" name=""unit_" & aiStayNum & "_" & aiThisReqmnt & """ " & "id=""" & strIptId & """ "
				If Not (IsVBPollingEnabled) Then
					' Not sure this onclick is even required without VB Polling.. (?)
					pO.Write "onclick=""BookingUI_UnitSelect(this);"" "
				End If
				pO.Write "value=""" & objUnit.UnitKey & """ "
				If (bSelected) Then
					pO.Write "checked=""checked"" "
				End If
				pO.Write "/>"
				pO.Write "<label for=""" & strIptId & """> "
				pO.Write objUnit.UnitName & " - " & BookingUI_NicePrice(mUnitStayTotal) & " " & asAvailClassName
				
				'if we have child pricing discount applied show the icon
				If bDiscountApplied Then
					pO.Write BookingUI_AvailClassIcon("DISCOUNT")
				End If
				
				pO.Write "</label>"
				pO.Write "</div>" & vbCrLf
				pO.Write "<div class=""pnPriceBase"">" & vbCrLf

'#MJ 29/04/2010 -	decision made not to show the price basis as the per person figure was always a guestimate, child pricing messes with the price so per person doesn't apply
'					also we now always deal with total stay prices
'				If bPerPerson Then
'					mPersonPerNight = mUnitPerNight/iNumPeople
'					pO.Write BookingUI_NicePrice(mPersonPerNight) & " " & Page.Resource("bookonline/unitselection/perpersonpernight", "per person per night") & ". "
'				Else
'					pO.Write BookingUI_NicePrice(mUnitPerNight) & " " & Page.Resource("bookonline/unitselection/perroomunitpernight", "per room/unit per night") & ". "
'				End If

				If iDaysBreakfast = iNumNights Then
					pO.Write Page.Resource("bookonline/unitselection/breakfastincluded","Breakfast included") & ". "
				ElseIf (iDaysBreakfast > 0) Then
					pO.Write Page.Resource("bookonline/unitselection/breakfastincludedon","Breakfast included on ") & iDaysBreakfast & " " & Page.Resource("bookonline/unitselection/day(s)","day(s)") & ". "
				End If

				If (iNumPeople < objUnit.MinOcc) Then
					If bPerPerson Then
						pO.Write "<br />"
						pO.Write Page.Resource("bookonline/unitselection/priceperpersonincludes", "Price Per Person includes") & " "
						pO.Write BookingUI_NicePrice(mPersonPerNight - UnitCostPerPerson / iNumNights)
						pO.Write Page.Resource("bookonline/unitselection/minimumoccupancysupplement", " minimum occupancy supplement") & ". "
					Else
						pO.Write Page.Resource("bookonline/unitselection/minoccupancyof","Min. occupancy of") & " " & objUnit.MinOcc & ". "
					End If
				End If
				pO.Write "<div class=""pnLinkedUnit"">" & BookingUI_LinkedUnitDesc(objUnit, objAllUnits) & "</div>"

				pO.Write "</div>" & vbCrLf

				If Not (objAvailEntry.IsLocal) Then
					pO.Write "<input type=""hidden"" "
					pO.Write "name=""uxml_" & aiStayNum & "_" & aiThisReqmnt & "_" & objUnit.UnitKey & """ "
					pO.Write "value=""" & Server.HtmlEncode(objUnit.EviivoMetaData) & """ />"
				End If
				
				If bRenderMaximumUnitsAvailable Then
					pO.Write "<div class=""maxAvailUnits"">"
					pO.Write "<p>"
					pO.Write "<span class=""maxAvailUnitsLabelPrefix"">" & Page.Resource("bookonline/unitselection/maxiumunitsavailableprefix", "Only ") &"</span>"
					pO.Write "<span class=""maxAvailUnitsValue"">" & objUnit.MaximumQuantityAvailable & "</span>"
					pO.Write "<span class=""maxAvailUnitsLabelSuffix"">" & Page.Resource("bookonline/unitselection/maxiumunitsavailablesuffix", " Rooms Remaining") &"</span>"
					pO.Write "</p>"
					pO.Write "</div>"
				End If
				
			End Function

			' ====================================================================================================
			' RENDER: Generate markup for booking buttons (only used by Acco, not Ticketing)
			' ====================================================================================================
			' SUMMARY: render BOOK and BACK buttons
			' [aiStayNum]: integer stay number [1-based]
			' [abPrecise]: boolean precise match (ie. hide BACK button)
			' <retval>: string output
			Private Function BookingUI_RenderButtons(ByVal aiStayNum, ByVal pO, ByVal bExternal)
				Dim strClass : strClass = "btnBookStay"

				If bExternal Then
					strClass = strClass & " redirect"
				End If
				
				pO.Write "<div class=""pnStayButtons"">"
				pO.Write "<input "
				pO.Write  "type=""image"" "
				pO.Write  "class=""" & strClass & """ "
				pO.Write  "name=""bookstay_" & aiStayNum & """ "

				' Not using ids with VB Polling layout
				If Not (IsVBPollingEnabled) Then
					pO.Write "id=""bookstay_" & aiStayNum & """ "
				End If

				pO.Write  "value=""" & Page.Resource("bookonline/btn/book","Book") & """ "
				pO.Write  "src=""" & Page.ImageResource("bookonline/btn/book", Context.ImageDir&"booking/book.gif") & """ "
				pO.Write  "alt=""" & Page.Resource("bookonline/btn/book","Book") & """ />"
				
				pO.Write "</div>" & vbCrLf

			End Function

			' ====================================================================================================
			' RENDER: Translate availability options -> action description string
			'  eg. Telephone Booking       -> "Submit a Booking Enquiry"
			'      Indicative Availability -> "Confirm availability"
			' ====================================================================================================
			' SUMMARY: describe the significant availability type (eg. Indicative, etc)
			' [abIndicative]: boolean indicates whether there is INDICATIVE availability
			' [abIndicValid]: boolean is INDICATIVE availability valid in this case
			' [abTeleBook]: boolean on telebook channel
			Function BookingUI_AvailClassName(ByVal abIndicative, ByVal abIndicValid, ByVal abTeleBook)
				' If telephone booking, there's only one option
				If abTeleBook Then
					BookingUI_AvailClassName = BookingUI_AvailClassIcon("TELE")
					Exit Function
				End If

				' If not telephone and not indicative, must be allocated
				If Not abIndicative Then
					BookingUI_AvailClassName = BookingUI_AvailClassIcon("ALLOC")
					Exit Function
				End If

				' Otherwise, get appropriate indicative option
				If abIndicValid Then
					BookingUI_AvailClassName = BookingUI_AvailClassIcon("INDIC")
				Else
					BookingUI_AvailClassName = BookingUI_AvailClassIcon("TELE")
				End If
			End Function

			' ====================================================================================================
			' RENDER:Translate avail class ID -> action description string
			'  eg. "ALLOC" -> "Book Online"
			'      "INDIC" -> "Confirm Availability"
			' ====================================================================================================
			' SUMMARY: get HTML for rendering the icon describing the availability class
			' [asAvailClassId]: string availability class ID (ie. ALLOC, INDIC or TELE)
			' <retval>: string image describing availClass
			Private	Function BookingUI_AvailClassIcon(asAvailClassId)
				Dim sIcon,sTxt, sImg
				
				'Select Case asAvailClassId
				'	Case "ALLOC"
				'		sIcon="icon_availClass_alloc"
				'		sTxt=Page.Resource("bookonline/btn/bookonline","Book Online")
				'	Case "INDIC"
				'		sIcon="icon_availClass_indic"
				'		sTxt=Page.Resource("bookonline/btn/confirmavailability","Confirm Availability")
				'	Case "DISCOUNT"
				'		sIcon="icon_availClass_discount"
				'		sTxt=Page.Resource("bookonline/btn/discountapplied","Child Pricing Discount Applied")
				'	Case Else
				'		sIcon="icon_availClass_tele"
				'		sTxt=Page.Resource("bookonline/btn/submitbookingenquiry","Submit a Booking Enquiry")
				'End Select
				
				sImg = Page.ImageResource("bookonline/icons/" & sIcon, Context.ImageDir&"booking/"&sIcon&".gif")
				BookingUI_AvailClassIcon="<img src="""& sImg & """ style=""vertical-align:middle;"" alt="""&sTxt&""" />"
			End Function

			' ====================================================================================================
			' RENDER: Format currency value
			' ====================================================================================================
			' SUMMARY: save space - only display price with pennies digits when fractional pounds
			' [amPrice]: money price to render
			' <retval>: string price
			Private Function BookingUI_NicePrice(amPrice)
				' Get price:
				' - MakePrice will also handle any currency conversion)
				' - MakePrice will apply an appropriate currency symbol
				Dim strPrice
				strPrice = Page.Functions.Money.MakePrice(amPrice)

				' If there's a trailing ".00" then trim it off
				' NB: Pretty sure we'll never get a price of the form "?.00" - it should always
				'     be "?0.00", but just in case check that we've got a suitable long string
				If (Len(strPrice) > 4) Then
					If (Right(strPrice, 3) = ".00") Then
						strPrice = Left(strPrice, Len(strPrice) - 3)
					End If
				End If

				' Return string ready for display
				BookingUI_NicePrice = Server.HTMLEncode(strPrice)
			End Function

			' ====================================================================================================
			' RENDER: Pull description of linked unit (includes name of linked unit, name of source unit and
			' size of linked unit)
			' ====================================================================================================
			' SUMMARY: get description of linked unit - this is the PHYSICAL unit description
			Private Function BookingUI_LinkedUnitDesc(ByVal objUnit, ByVal objAllUnits)
				
				' If either UnitName of LinkedUnitName absent, return blank
				Dim sUnitName: sUnitName = objUnit.UnitName
				Dim sLinkedUnitName: sLinkedUnitName = objUnit.LinkUnitName
				If (IsNull(sUnitName) Or sUnitName = "" Or IsNull(sLinkedUnitName) Or sLinkedUnitName = "") Then
					BookingUI_LinkedUnitDesc = ""
					Exit Function
				End If

				' 2014-08-26 DWR: We need to retrieve the capacity of the unit that this linked unit is linked to. This data is not available in the avail
				' data from TOv2 since it is not included in the data from the Availability Component. It is why the "all units" data must be passed into
				' this method. This change addresses FogBugz 12998.
				Dim objParentUnit: Set objParentUnit = Nothing
				Dim intIndex: For intIndex = 0 To (objAllUnits.Count - 1)
					If (objAllUnits.getItem(intIndex).Key = objUnit.LinkUnitKey) Then
						Set objParentUnit = objAllUnits.getItem(intIndex)
						Exit For
					End If
				Next
				If (objParentUnit is Nothing) Then
					Page.PrintTraceWarning "Unable to locate parent unit (" & objUnit.LinkUnitKey & ") for linked unit " & objUnit.UnitKey
					BookingUI_LinkedUnitDesc = ""
					Exit Function
				End If

				BookingUI_LinkedUnitDesc = _
					Replace( _
						Replace( _
							Replace( _
								Page.Resource( _
									"bookonline/unitselection/alsosoldaswithpersoncapacity", _
									"(<i>#linkedunitname#</i> sold as #unitname# with #linkunitsize# person capacity)"), _
								"#linkedunitname#", _
								sLinkedUnitName), _
							"#unitname#", _
							sUnitName), _
						"#linkunitsize#", _
						objParentUnit.Capacity)
			End Function


			' ====================================================================================================
			' RENDER: This handles all of the rendering for ticketing - none of the StaySummary, StayDetails,
			' RenderButtons malarkey is required
			' ====================================================================================================
			Private Function BookingUI_TicketsSummary(objAvailEntry,adtStartNight,pO)
				Dim iTotal,iSubTotal,iSelectedQty, intIndexUnit, objUnit, strPriceBasis
		
				If objAvailEntry.Units.Count > 0 Then
					pO.Write "<div id=""availabilityCalendarTableWrapper"">"
					pO.Write "<h3>" & Page.Resource("bookonline/unitselection/ticketsavailable", "Tickets Available:") & "</h3>"
					pO.Write "<table id=""availabilityCalendarTable"" summary="""&Page.Resource("bookonline/unitselection/ticketsavailable","Tickets Available")&""" border=""1"">"
					pO.Write "<thead>"
					pO.Write "<tr class=""heading"">"
					pO.Write "<th class=""unit"">"&Page.Resource("bookonline/unitselection/tickets","Tickets")&"</th>"
					pO.Write "<th class=""select"">"&Page.Resource("bookonline/unitselection/selection","Selection")&"</th>"
					pO.Write "<th class=""date"">"&Page.Resource("bookonline/unitselection/date","Date")&"</th>"
					pO.Write "<th class=""total"">"&Page.Resource("bookonline/unitselection/total","Total")&"</th>"
					pO.Write "</tr>"
					pO.Write "<tr>"
					pO.Write "<th></th>"
					pO.Write "<th class=""number"">"&Page.Resource("bookonline/unitselection/nooftickets","No.Tickets")&"</th>"
					pO.Write "<th class=""staydate"">"&Page.Functions.Dates.NiceDateGuts(adtStartNight,True,True)&"</th>"
					pO.Write "<th class=""total""></th>"
					pO.Write "</tr>"
					pO.Write "</thead>"
					pO.Write "<tbody>"
					iTotal = 0

					For intIndexUnit = 0 To objAvailEntry.Units.Count - 1
						Set objUnit = objAvailEntry.Units.GetItem(intIndexUnit)
						
						iSelectedQty = CLng(Request.Form("unit_" & objUnit.UnitKey))
						
						If objUnit.PerPerson Then 
							strPriceBasis = "per per"
						Else
							strPriceBasis = "per tic"
						End If 
					
						pO.Write "<tr id=""row_"&objUnit.UnitKey&""">"
						pO.Write "<td class=""unit"">"&objUnit.UnitName&"</td>"
						pO.Write "<td class=""select"">"&Page.Functions.Booking.DrawSelectRange("unit_" &objUnit.UnitKey, 0, objUnit.UnitCount, iSelectedQty)&"</td>"
						pO.Write "<td class=""price"">" & Server.HTMLEncode(Page.Functions.Money.MakePrice(objUnit.StayTotalPayable))&"</td>"
						pO.Write _
							"<td class=""total"">" & _
								"<input type=""hidden"" name=""data_"&objUnit.UnitKey&""" id=""data_"&objUnit.UnitKey&""" value="""&objUnit.UnitCount&","&objUnit.MinOcc&","&objUnit.UnitSize&","&strPriceBasis&","&objUnit.StayTotalPayable&""">" & _
								Server.HTMLEncode(Page.Functions.Money.MakePrice(objUnit.StayTotalPayable * iSelectedQty)) & _
							"</td>"
						pO.Write "</tr>"
						iTotal = iTotal + objUnit.StayTotalPayable * iSelectedQty
						
				Next
					iSubTotal = iSubTotal + iTotal

					pO.Write "</tbody>"
					pO.Write "</table>"
					pO.Write "</div>"
					pO.Write "<table id=""availabilityTotals"" summary=""Totals"" border=""1"">"
					pO.Write "<tr>"
					pO.Write "<th>"&Page.Resource("bookonline/unitselection/grandtotal","Grand Total")&"</th>"
					pO.Write "<noscript>"
					pO.Write "<td><input type=""image"" src="""&Page.ImageResource("bookonline/unitselection/recalculate", Context.ImageDir & "booking/bookrecalculate.gif") & """ name=""recalculate"" value=""recalculate"" class=""submit""/></td>"
					pO.Write "</noscript>"
					pO.Write "<td id=""AvCalTotal"">"&Server.HTMLEncode(Page.Functions.Money.MakePrice(iSubTotal))&"</td>"
					pO.Write "</tr>"
					pO.Write "</table>"
					pO.Write "<input type=""image"" src="""&Page.ImageResource("bookonline/btn/bookticketing" ,Context.ImageDir & "booking/bookticketing.gif") & """ name=""bookit"" value="""&Page.Resource("bookonline/btn/book","Book")&""" alt="""&Page.Resource("bookonline/btn/book","Book")&""" class=""submit""/>"
				End If
			End Function


			' ====================================================================================================
			' RENDER: If Site Param Booking_ForceExternal is set, then all  bookings are forced to external sites
			' (the site depends upon the product's Estate) - this function gets the external destination
			' 2009-02-19 DWR: This handles Local availability products, FrontDesk products will be treated like
			' this as well if VB Polling is disabled, if it is ENabled then the FrontDesk products will act like
			' any other external supplier and should deep-link into their site.
			' ====================================================================================================
			Private Function GetExtBookUrlFromProductEstate(asEstateID)
				Dim strPostUrl_Ext, strPostUrl_ExtDflt, aryExtBookEstate,i
					' 2009-02-13 DWR: Can't remove spaces from content here because estate ids can have
					' spaces in (eg. "Arun DC" in TSE)
					aryExtBookEstate = Split(Replace(Trim("" & Page.Site.Params("Booking_ExtBookEstateMapping")),vbCrLf,""),",")
					For i=0 To UBound(aryExtBookEstate)-1 Step 2
  						If UCase(Trim(aryExtBookEstate(i))) = "DEFAULT" Then
	  						strPostUrl_ExtDflt = aryExtBookEstate(i+1)
  						ElseIf UCase(Trim(aryExtBookEstate(i))) = UCase(Trim(asEstateID)) Then
  							strPostUrl_Ext = aryExtBookEstate(i+1)
  							Exit For
	  					End If
					Next

					If strPostUrl_Ext <> "" Then
						GetExtBookUrlFromProductEstate = strPostUrl_Ext
						Page.PrintTrace "GetExtBookUrlFromProductEstate: Product Estate ID = " & asEstateID & ", External Book Url = " & strPostUrl_Ext
					ElseIf strPostUrl_ExtDflt <> "" Then
						GetExtBookUrlFromProductEstate = strPostUrl_ExtDflt
						Page.PrintTrace "GetExtBookUrlFromProductEstate: Product Estate ID = " & asEstateID & ", Using Default External Book Url = " & strPostUrl_ExtDflt
					Else
						Err.Raise vbObjectError + 1, "ETWP.Booking_UnitSelection Control", "Failed to get External Booking Url [" & asEstateID & "]"
					End If
					
			
			End Function
			
			Private Function InitExternalBookingSettings()
				If Page.Site.Params("Booking_ForceExternal") And Trim(""&Page.Site.Params("Booking_ExtBookEstateMapping")) <> "" Then
					IsExternalBooking = True
				Else
					IsExternalBooking = False
				End If
			End Function

			' ====================================================================================================
			' MISC: Since the RenderSettings.BookingRequirement references passed into here are usually read-only
			' instances from the Page.Functions.GetSharedObject method, we'll need to make a local copy that we
			' can manipulate (since in some cases we need to mess about with the values)
			' ====================================================================================================
			Private Function GetEditableBookingRequirement(ByVal objBookingRequirement)

				Dim objBookingRequirementNew: Set objBookingRequirementNew = Page.Functions.GetNewObject("BookingRequirement")
				With objBookingRequirementNew
					.VisitDate = objBookingRequirement.VisitDate
					.Nights = objBookingRequirement.Nights
					.FlexibleRange = objBookingRequirement.FlexibleRange
					.Adults = objBookingRequirement.Adults
					.Children = objBookingRequirement.Children
					.ChildAges = objBookingRequirement.ChildAges
					.IsEviivoBooking = objBookingRequirement.IsEviivoBooking
					.Consumer = objBookingRequirement.Consumer
					.Offer = objBookingRequirement.Offer
					.BookingPassword = objBookingRequirement.BookingPassword
					.Product = objBookingRequirement.Product
					.Requirement = objBookingRequirement.Requirement
					.RequirementRef = objBookingRequirement.RequirementRef
					' NP 2012-03-12: RoomRequirements are needed
					' See GenerateRequirementFormData and Page.Functions.Booking.GenerateRequirementKeyValueData
					' the "NumRoomReq" value is part of the RoomRequirement, if it is not available then GenerateRequirementKeyValueData
					' sets default values for the adult and number of room requirements both to 1.
					' Requirements are not being passed to the RequirementSummary control correctly because the BookingRequestDictionary
					' is being overwritten with these incorrect default values.
					.RoomRequirements = objBookingRequirement.RoomRequirements
				End With
				Set GetEditableBookingRequirement = objBookingRequirementNew
			End Function
