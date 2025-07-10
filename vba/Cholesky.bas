Option Explicit

' Compute the Cholesky decomposition of a symmetric positive definite matrix.
' The input A is expected to be a 2-D Variant array or a Range containing
' numeric values. The function returns a 2-D Variant array containing the
' lower triangular matrix L such that A = L * WorksheetFunction.Transpose(L).
Public Function CholeskyDecomposition(A As Variant) As Variant
    Dim n As Long, i As Long, j As Long, k As Long
    Dim L() As Double
    Dim sum As Double

    ' Convert input to array if it is a Range
    If TypeName(A) = "Range" Then
        A = A.Value
    End If

    n = UBound(A, 1)
    ReDim L(1 To n, 1 To n)

    For i = 1 To n
        For j = 1 To i
            sum = A(i, j)
            For k = 1 To j - 1
                sum = sum - L(i, k) * L(j, k)
            Next k

            If i = j Then
                If sum <= 0 Then
                    Err.Raise vbObjectError + 513, "CholeskyDecomposition", _
                              "Matrix is not positive definite"
                End If
                L(i, j) = Sqr(sum)
            Else
                L(i, j) = sum / L(j, j)
            End If
        Next j
    Next i

    CholeskyDecomposition = L
End Function

' Example usage demonstrating decomposition of a 3x3 matrix
Public Sub TestCholesky()
    Dim A(1 To 3, 1 To 3) As Double
    Dim L As Variant
    Dim i As Long, j As Long

    ' Define a symmetric positive definite matrix
    A(1, 1) = 4: A(1, 2) = 12: A(1, 3) = -16
    A(2, 1) = 12: A(2, 2) = 37: A(2, 3) = -43
    A(3, 1) = -16: A(3, 2) = -43: A(3, 3) = 98

    L = CholeskyDecomposition(A)

    ' Output the result to the Immediate Window
    For i = 1 To 3
        For j = 1 To 3
            Debug.Print L(i, j); " "
        Next j
        Debug.Print
    Next i
End Sub
