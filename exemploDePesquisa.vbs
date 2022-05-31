<code><pre><%

sub ev_onComplimentBefore(ev)' [6]
    Response.write("Event onComplimentBefore has been fired. I was really expecting this method to say: 'Hello World " & ev.Arguments.item("firstname") & " " & ev.Arguments.item("lastname") & " (" & ev.Arguments.item("nickname") & ")'" & vbNewline)
end sub

sub ev_onComplimentAfter(ev)
    Response.write("Event onComplimentAfter has been fired" & vbNewline)
end sub

dim CwE : set CwE = new ClassWithEvents
call CwE.onComplimentBefore.addHandler("ev_onComplimentBefore")' [7]
call CwE.onComplimentAfter.addHandler("ev_onComplimentAfter")
call CwE.compliment("Fabio", "Nagao", "nagaozen")
set CwE = nothing

%></pre></code>