
<%
function CPMF(k,n,p)
	dim i, result
	for i = 0 to (k-1)
		result = result + PMF(i,n,p)
	next
	CPMF = result
end function

function PMF(k,n,p)
	dim result
	result = ((factorial(n) / (factorial(k) * factorial(n-k))) * ((p^k)*((1-p)^(n-k))))
	PMF = result
end function

function factorial(x)
	dim result, i
	result = 1
	if x > 1 then
		for i = 2 to x
			result = result * i
		next
	end if
	factorial = result
end function
%>
