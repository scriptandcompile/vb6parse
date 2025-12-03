//! # Rnd Function
//!
//! Returns a Single containing a pseudo-random number.
//!
//! ## Syntax
//!
//! ```vb
//! Rnd[(number)]
//! ```
//!
//! ## Parameters
//!
//! - `number` - Optional. Single or any valid numeric expression.
//!
//! ## Return Value
//!
//! Returns a `Single` value greater than or equal to 0 and less than 1.
//!
//! The behavior depends on the `number` argument:
//!
//! | Number Value | Result |
//! |--------------|--------|
//! | < 0 | Same number every time, using `number` as the seed |
//! | > 0 | Next random number in sequence |
//! | = 0 | Most recently generated number |
//! | Omitted | Next random number in sequence |
//!
//! ## Remarks
//!
//! The `Rnd` function returns a pseudo-random number from a deterministic sequence. It is used for generating random values in games, simulations, testing, and sampling.
//!
//! Each new number is the product of the previous number times a constant (a) plus another constant (c).
//! The result is kept modulo a third number (m), which is invariably the width of its ‘long’ integer size (or fraction thereof).
//! Standard examples are m=2^64, 2^48 or 2^32.
//! Mathematically, this is represented as: r(i+1)=r(i)*a+c (mod m), where r(0) is the ‘seed’ value.
//!
//! Although the algorithm in use can be obtained from analyzing a string of outputs from the generator and deducing the values of a, c and M.
//! Microsoft has published their constants on their web site and it turns out that they use an LCPRNG with:
//!
//! ```a = 16598013, c = 2820163 and m=2^24 = 16777216```
//!
//! [Reference](http://www.noesis.net.au/main/Resources/Resources/prng_files/vba_prng.html)
//!
//! **Important Notes**:
//! - Returns values in range [0, 1) - includes 0, excludes 1
//! - Same seed produces same sequence (deterministic)
//! - Use `Randomize` statement to initialize random number generator with time-based seed
//! - Without `Randomize`, same sequence is generated each program run
//! - Passing negative number sets seed for reproducible sequences
//! - Passing 0 returns last generated number (useful for debugging)
//!
//! **Common Usage Pattern**:
//! ```vb
//! Randomize               ' Initialize with time-based seed
//! x = Rnd                 ' Get random number 0 <= x < 1
//! ```
//!
//! **Generating Random Integers**:
//! ```vb
//! ' Random integer from min to max (inclusive)
//! randomInt = Int((max - min + 1) * Rnd + min)
//! ```
//!
//! **Seeding for Reproducibility**:
//! ```vb
//! Rnd -1                  ' Reset to use seed
//! Randomize seed          ' Set specific seed
//! x = Rnd                 ' Get first number in sequence
//! ```
//!
//! ## Typical Uses
//!
//! 1. **Games**: Generate random positions, events, or outcomes
//! 2. **Simulations**: Create random data for Monte Carlo simulations
//! 3. **Testing**: Generate random test data
//! 4. **Sampling**: Random selection from datasets
//! 5. **Shuffling**: Randomize order of items
//! 6. **Password Generation**: Create random character sequences
//! 7. **Animation**: Random movement or effects
//! 8. **Data Masking**: Generate random replacement values
//!
//! ## Basic Examples
//!
//! ### Example 1: Simple Random Number
//! ```vb
//! ' Generate random number between 0 and 1
//! Randomize
//! Dim randomValue As Single
//! randomValue = Rnd()  ' e.g., 0.7234567
//! ```
//!
//! ### Example 2: Random Integer in Range
//! ```vb
//! ' Generate random integer from 1 to 100
//! Randomize
//! Dim randomInt As Integer
//! randomInt = Int(Rnd * 100) + 1
//! ```
//!
//! ### Example 3: Random Dice Roll
//! ```vb
//! ' Simulate rolling a six-sided die
//! Randomize
//! Dim diceRoll As Integer
//! diceRoll = Int(Rnd * 6) + 1  ' Returns 1-6
//! ```
//!
//! ### Example 4: Random Selection
//! ```vb
//! ' Select random item from array
//! Randomize
//! Dim items(5) As String
//! Dim randomIndex As Integer
//!
//! items(0) = "Apple"
//! items(1) = "Banana"
//! items(2) = "Cherry"
//! items(3) = "Date"
//! items(4) = "Elderberry"
//!
//! randomIndex = Int(Rnd * 5)
//! MsgBox "Selected: " & items(randomIndex)
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `RandomInteger`
//! ```vb
//! Function RandomInteger(minValue As Long, maxValue As Long) As Long
//!     ' Generate random integer in range [minValue, maxValue]
//!     RandomInteger = Int((maxValue - minValue + 1) * Rnd + minValue)
//! End Function
//! ```
//!
//! ### Pattern 2: `RandomDouble`
//! ```vb
//! Function RandomDouble(minValue As Double, maxValue As Double) As Double
//!     ' Generate random double in range [minValue, maxValue)
//!     RandomDouble = (maxValue - minValue) * Rnd + minValue
//! End Function
//! ```
//!
//! ### Pattern 3: `RandomBoolean`
//! ```vb
//! Function RandomBoolean() As Boolean
//!     ' Generate random True/False
//!     RandomBoolean = (Rnd >= 0.5)
//! End Function
//! ```
//!
//! ### Pattern 4: `RandomChoice`
//! ```vb
//! Function RandomChoice(items As Variant) As Variant
//!     ' Select random item from array
//!     Dim index As Integer
//!     
//!     If Not IsArray(items) Then
//!         Err.Raise 5, , "Parameter must be an array"
//!     End If
//!     
//!     index = Int(Rnd * (UBound(items) - LBound(items) + 1)) + LBound(items)
//!     RandomChoice = items(index)
//! End Function
//! ```
//!
//! ### Pattern 5: `ShuffleArray`
//! ```vb
//! Sub ShuffleArray(arr As Variant)
//!     ' Randomly shuffle array elements (Fisher-Yates algorithm)
//!     Dim i As Integer
//!     Dim j As Integer
//!     Dim temp As Variant
//!     
//!     For i = UBound(arr) To LBound(arr) + 1 Step -1
//!         j = Int(Rnd * (i - LBound(arr) + 1)) + LBound(arr)
//!         
//!         temp = arr(i)
//!         arr(i) = arr(j)
//!         arr(j) = temp
//!     Next i
//! End Sub
//! ```
//!
//! ### Pattern 6: `RandomString`
//! ```vb
//! Function RandomString(length As Integer, _
//!                      Optional chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789") As String
//!     ' Generate random string of specified length
//!     Dim i As Integer
//!     Dim result As String
//!     Dim index As Integer
//!     
//!     result = ""
//!     
//!     For i = 1 To length
//!         index = Int(Rnd * Len(chars)) + 1
//!         result = result & Mid(chars, index, 1)
//!     Next i
//!     
//!     RandomString = result
//! End Function
//! ```
//!
//! ### Pattern 7: `RandomColor`
//! ```vb
//! Function RandomColor() As Long
//!     ' Generate random RGB color
//!     Dim r As Integer
//!     Dim g As Integer
//!     Dim b As Integer
//!     
//!     r = Int(Rnd * 256)
//!     g = Int(Rnd * 256)
//!     b = Int(Rnd * 256)
//!     
//!     RandomColor = RGB(r, g, b)
//! End Function
//! ```
//!
//! ### Pattern 8: `RandomPercentage`
//! ```vb
//! Function RandomPercentage(probability As Double) As Boolean
//!     ' Return True with specified probability (0.0 to 1.0)
//!     RandomPercentage = (Rnd < probability)
//! End Function
//! ```
//!
//! ### Pattern 9: `RandomDate`
//! ```vb
//! Function RandomDate(startDate As Date, endDate As Date) As Date
//!     ' Generate random date between start and end dates
//!     Dim daysDiff As Long
//!     Dim randomDays As Long
//!     
//!     daysDiff = DateDiff("d", startDate, endDate)
//!     randomDays = Int(Rnd * (daysDiff + 1))
//!     
//!     RandomDate = DateAdd("d", randomDays, startDate)
//! End Function
//! ```
//!
//! ### Pattern 10: `WeightedRandom`
//! ```vb
//! Function WeightedRandom(weights() As Double) As Integer
//!     ' Select index based on weighted probabilities
//!     Dim total As Double
//!     Dim i As Integer
//!     Dim randomValue As Double
//!     Dim cumulative As Double
//!     
//!     ' Calculate total weight
//!     total = 0
//!     For i = LBound(weights) To UBound(weights)
//!         total = total + weights(i)
//!     Next i
//!     
//!     ' Generate random value
//!     randomValue = Rnd * total
//!     
//!     ' Find corresponding index
//!     cumulative = 0
//!     For i = LBound(weights) To UBound(weights)
//!         cumulative = cumulative + weights(i)
//!         If randomValue < cumulative Then
//!             WeightedRandom = i
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     WeightedRandom = UBound(weights)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Random Number Generator Class
//! ```vb
//! ' Comprehensive random number generator with multiple distributions
//! Class RandomGenerator
//!     Private m_seed As Long
//!     Private m_seeded As Boolean
//!     
//!     Public Sub Initialize(Optional seed As Long = 0)
//!         If seed <> 0 Then
//!             m_seed = seed
//!             Rnd -1
//!             Randomize seed
//!             m_seeded = True
//!         Else
//!             Randomize
//!             m_seeded = False
//!         End If
//!     End Sub
//!     
//!     Public Function NextDouble() As Double
//!         ' Get next random double [0, 1)
//!         NextDouble = Rnd
//!     End Function
//!     
//!     Public Function NextInteger(minValue As Long, maxValue As Long) As Long
//!         ' Get random integer [minValue, maxValue]
//!         If minValue > maxValue Then
//!             Err.Raise 5, , "minValue must be <= maxValue"
//!         End If
//!         
//!         NextInteger = Int((maxValue - minValue + 1) * Rnd + minValue)
//!     End Function
//!     
//!     Public Function NextBoolean() As Boolean
//!         ' Get random boolean
//!         NextBoolean = (Rnd >= 0.5)
//!     End Function
//!     
//!     Public Function NextGaussian(Optional mean As Double = 0, _
//!                                 Optional stdDev As Double = 1) As Double
//!         ' Generate Gaussian (normal) distribution using Box-Muller transform
//!         Dim u1 As Double, u2 As Double
//!         Dim z0 As Double
//!         
//!         u1 = Rnd
//!         u2 = Rnd
//!         
//!         ' Box-Muller transform
//!         z0 = Sqr(-2 * Log(u1)) * Cos(2 * 3.14159265358979 * u2)
//!         
//!         NextGaussian = mean + stdDev * z0
//!     End Function
//!     
//!     Public Function NextBytes(count As Long) As Byte()
//!         ' Generate array of random bytes
//!         Dim bytes() As Byte
//!         Dim i As Long
//!         
//!         ReDim bytes(0 To count - 1)
//!         
//!         For i = 0 To count - 1
//!             bytes(i) = Int(Rnd * 256)
//!         Next i
//!         
//!         NextBytes = bytes
//!     End Function
//!     
//!     Public Sub Reset()
//!         ' Reset to original seed if seeded
//!         If m_seeded Then
//!             Rnd -1
//!             Randomize m_seed
//!         Else
//!             Randomize
//!         End If
//!     End Sub
//!     
//!     Public Function GetSeed() As Long
//!         GetSeed = m_seed
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Monte Carlo Simulator
//! ```vb
//! ' Monte Carlo simulation for estimating probabilities
//! Module MonteCarloSimulator
//!     Public Function EstimatePi(iterations As Long) As Double
//!         ' Estimate Pi using Monte Carlo method
//!         Dim insideCircle As Long
//!         Dim i As Long
//!         Dim x As Double, y As Double
//!         
//!         Randomize
//!         insideCircle = 0
//!         
//!         For i = 1 To iterations
//!             x = Rnd
//!             y = Rnd
//!             
//!             If (x * x + y * y) < 1 Then
//!                 insideCircle = insideCircle + 1
//!             End If
//!         Next i
//!         
//!         EstimatePi = 4 * (insideCircle / iterations)
//!     End Function
//!     
//!     Public Function SimulateDiceRolls(numDice As Integer, _
//!                                      numRolls As Long) As Long()
//!         ' Simulate rolling multiple dice and return frequency distribution
//!         Dim results() As Long
//!         Dim minSum As Integer, maxSum As Integer
//!         Dim i As Long, j As Integer
//!         Dim sum As Integer
//!         
//!         minSum = numDice
//!         maxSum = numDice * 6
//!         
//!         ReDim results(minSum To maxSum)
//!         
//!         Randomize
//!         
//!         For i = 1 To numRolls
//!             sum = 0
//!             For j = 1 To numDice
//!                 sum = sum + Int(Rnd * 6) + 1
//!             Next j
//!             results(sum) = results(sum) + 1
//!         Next i
//!         
//!         SimulateDiceRolls = results
//!     End Function
//!     
//!     Public Function SimulateStockPrice(startPrice As Double, _
//!                                       days As Long, _
//!                                       volatility As Double, _
//!                                       drift As Double) As Double()
//!         ' Simulate stock price using geometric Brownian motion
//!         Dim prices() As Double
//!         Dim i As Long
//!         Dim randomShock As Double
//!         
//!         ReDim prices(0 To days)
//!         prices(0) = startPrice
//!         
//!         Randomize
//!         
//!         For i = 1 To days
//!             randomShock = (Rnd - 0.5) * 2  ' -1 to 1
//!             prices(i) = prices(i - 1) * (1 + drift + volatility * randomShock)
//!         Next i
//!         
//!         SimulateStockPrice = prices
//!     End Function
//! End Module
//! ```
//!
//! ### Example 3: Random Data Generator
//! ```vb
//! ' Generate random test data for various purposes
//! Class RandomDataGenerator
//!     Private m_firstNames() As String
//!     Private m_lastNames() As String
//!     Private m_emailDomains() As String
//!     
//!     Public Sub Initialize()
//!         Randomize
//!         
//!         m_firstNames = Array("John", "Jane", "Michael", "Sarah", "David", _
//!                             "Emily", "James", "Emma", "Robert", "Lisa")
//!         m_lastNames = Array("Smith", "Johnson", "Williams", "Brown", "Jones", _
//!                            "Garcia", "Miller", "Davis", "Rodriguez", "Martinez")
//!         m_emailDomains = Array("gmail.com", "yahoo.com", "hotmail.com", _
//!                               "outlook.com", "example.com")
//!     End Sub
//!     
//!     Public Function GenerateName() As String
//!         ' Generate random full name
//!         Dim firstName As String
//!         Dim lastName As String
//!         
//!         firstName = m_firstNames(Int(Rnd * (UBound(m_firstNames) + 1)))
//!         lastName = m_lastNames(Int(Rnd * (UBound(m_lastNames) + 1)))
//!         
//!         GenerateName = firstName & " " & lastName
//!     End Function
//!     
//!     Public Function GenerateEmail(Optional name As String = "") As String
//!         ' Generate random email address
//!         Dim localPart As String
//!         Dim domain As String
//!         
//!         If name = "" Then
//!             localPart = "user" & Int(Rnd * 10000)
//!         Else
//!             localPart = LCase(Replace(name, " ", "."))
//!         End If
//!         
//!         domain = m_emailDomains(Int(Rnd * (UBound(m_emailDomains) + 1)))
//!         
//!         GenerateEmail = localPart & "@" & domain
//!     End Function
//!     
//!     Public Function GeneratePhoneNumber() As String
//!         ' Generate random US phone number
//!         Dim areaCode As String
//!         Dim exchange As String
//!         Dim number As String
//!         
//!         areaCode = Format(Int(Rnd * 900) + 100, "000")
//!         exchange = Format(Int(Rnd * 900) + 100, "000")
//!         number = Format(Int(Rnd * 10000), "0000")
//!         
//!         GeneratePhoneNumber = "(" & areaCode & ") " & exchange & "-" & number
//!     End Function
//!     
//!     Public Function GenerateAddress() As String
//!         ' Generate random street address
//!         Dim streetNumber As String
//!         Dim streetNames() As String
//!         Dim streetName As String
//!         Dim streetTypes() As String
//!         Dim streetType As String
//!         
//!         streetNames = Array("Main", "Oak", "Maple", "Cedar", "Elm", _
//!                           "Washington", "Park", "Lake", "Hill", "Pine")
//!         streetTypes = Array("St", "Ave", "Blvd", "Rd", "Ln", "Dr")
//!         
//!         streetNumber = Int(Rnd * 9900) + 100
//!         streetName = streetNames(Int(Rnd * (UBound(streetNames) + 1)))
//!         streetType = streetTypes(Int(Rnd * (UBound(streetTypes) + 1)))
//!         
//!         GenerateAddress = streetNumber & " " & streetName & " " & streetType
//!     End Function
//!     
//!     Public Function GeneratePassword(length As Integer) As String
//!         ' Generate random password with mixed characters
//!         Dim chars As String
//!         Dim i As Integer
//!         Dim result As String
//!         Dim index As Integer
//!         
//!         chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*"
//!         result = ""
//!         
//!         For i = 1 To length
//!             index = Int(Rnd * Len(chars)) + 1
//!             result = result & Mid(chars, index, 1)
//!         Next i
//!         
//!         GeneratePassword = result
//!     End Function
//! End Class
//! ```
//!
//! ### Example 4: Card Deck Shuffler
//! ```vb
//! ' Simulate a deck of cards with shuffling
//! Class CardDeck
//!     Private Type Card
//!         Suit As String
//!         Rank As String
//!         Value As Integer
//!     End Type
//!     
//!     Private m_cards() As Card
//!     Private m_currentCard As Integer
//!     
//!     Public Sub Initialize()
//!         Dim suits() As String
//!         Dim ranks() As String
//!         Dim i As Integer, j As Integer
//!         Dim cardIndex As Integer
//!         
//!         suits = Array("Hearts", "Diamonds", "Clubs", "Spades")
//!         ranks = Array("2", "3", "4", "5", "6", "7", "8", "9", "10", _
//!                      "Jack", "Queen", "King", "Ace")
//!         
//!         ReDim m_cards(0 To 51)
//!         cardIndex = 0
//!         
//!         For i = 0 To 3
//!             For j = 0 To 12
//!                 m_cards(cardIndex).Suit = suits(i)
//!                 m_cards(cardIndex).Rank = ranks(j)
//!                 m_cards(cardIndex).Value = j + 2
//!                 cardIndex = cardIndex + 1
//!             Next j
//!         Next i
//!         
//!         m_currentCard = 0
//!     End Sub
//!     
//!     Public Sub Shuffle()
//!         ' Shuffle deck using Fisher-Yates algorithm
//!         Dim i As Integer
//!         Dim j As Integer
//!         Dim temp As Card
//!         
//!         Randomize
//!         
//!         For i = 51 To 1 Step -1
//!             j = Int(Rnd * (i + 1))
//!             
//!             temp = m_cards(i)
//!             m_cards(i) = m_cards(j)
//!             m_cards(j) = temp
//!         Next i
//!         
//!         m_currentCard = 0
//!     End Sub
//!     
//!     Public Function DrawCard() As String
//!         ' Draw next card from deck
//!         Dim card As Card
//!         
//!         If m_currentCard > 51 Then
//!             DrawCard = "No cards left"
//!             Exit Function
//!         End If
//!         
//!         card = m_cards(m_currentCard)
//!         m_currentCard = m_currentCard + 1
//!         
//!         DrawCard = card.Rank & " of " & card.Suit
//!     End Function
//!     
//!     Public Function CardsRemaining() As Integer
//!         CardsRemaining = 52 - m_currentCard
//!     End Function
//!     
//!     Public Sub Reset()
//!         m_currentCard = 0
//!     End Sub
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `Rnd` function rarely generates errors, but there are some considerations:
//!
//! **Type Mismatch (Error 13)**:
//! - Occurs if the optional number parameter cannot be converted to numeric type
//!
//! Example error handling:
//!
//! ```vb
//! On Error Resume Next
//! Dim randomValue As Single
//! randomValue = Rnd(userInput)
//! If Err.Number <> 0 Then
//!     MsgBox "Invalid seed value"
//!     randomValue = Rnd  ' Use default behavior
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - `Rnd` is very fast - can generate millions of numbers per second
//! - Not cryptographically secure - don't use for security purposes
//! - Same seed produces same sequence (deterministic)
//! - For better randomness, call `Randomize` at program start
//! - Calling `Randomize` frequently can reduce randomness quality
//! - Consider caching random values if generating many at once
//!
//! ## Best Practices
//!
//! 1. **Always Randomize**: Call `Randomize` once at program start for non-deterministic behavior
//! 2. **Use Seed for Testing**: Use fixed seed (negative number) for reproducible test cases
//! 3. **Validate Ranges**: Check min/max values when generating random integers
//! 4. **Avoid Modulo Bias**: Use `Int(Rnd * n)` not `Rnd Mod n` for uniform distribution
//! 5. **Don't Use for Security**: Not suitable for passwords, encryption keys, or security tokens
//! 6. **Cache Rnd Values**: If generating many values, avoid repeated function calls overhead
//! 7. **Document Seed Usage**: Clearly document when using fixed seeds for reproducibility
//! 8. **Test Edge Cases**: Verify behavior at range boundaries (min, max values)
//! 9. **Use Helper Functions**: Wrap Rnd in helper functions for cleaner code
//! 10. **Consider Distribution**: Understand that Rnd provides uniform distribution
//!
//! ## Comparison with Related Functions
//!
//! | Function/Statement | Purpose | Returns | Use Case |
//! |-------------------|---------|---------|----------|
//! | **Rnd** | Random number | Single [0, 1) | Generate random values |
//! | **Randomize** | Initialize RNG | Nothing (statement) | Seed random number generator |
//! | **Int** | Integer part | Integer | Convert Rnd to integer range |
//! | **Timer** | Elapsed seconds | Single | Often used with Randomize |
//!
//! ## Platform and Version Notes
//!
//! - Available in all versions of VB6 and VBA
//! - Uses linear congruential generator (LCG) algorithm
//! - Not cryptographically secure
//! - Sequence repeats after approximately 16 million numbers
//! - In VB.NET, replaced by `System.Random` class
//! - `Randomize` uses `Timer` function by default if no seed provided
//!
//! ## Limitations
//!
//! - Not cryptographically secure (predictable sequence)
//! - Limited period (sequence repeats after ~16M values)
//! - No built-in support for other distributions (normal, exponential, etc.)
//! - Cannot generate random integers directly (requires Int/Floor conversion)
//! - Thread safety not guaranteed in multi-threaded scenarios
//! - Quality lower than modern RNGs (Mersenne Twister, etc.)
//!
//! ## Related Functions
//!
//! - `Randomize`: Initializes the random number generator with a seed
//! - `Int`: Returns the integer portion of a number (used to convert Rnd to integer range)
//! - `Timer`: Returns seconds since midnight (often used with Randomize)

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn test_rnd_basic() {
        let source = r#"
Dim randomValue As Single
randomValue = Rnd()
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_no_parens() {
        let source = r#"
Dim x As Single
x = Rnd
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_with_argument() {
        let source = r#"
Dim result As Single
result = Rnd(-1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_if_statement() {
        let source = r#"
If Rnd > 0.5 Then
    MsgBox "Heads"
Else
    MsgBox "Tails"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_function_return() {
        let source = r#"
Function GetRandomValue() As Single
    GetRandomValue = Rnd
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_variable_assignment() {
        let source = r#"
Dim randomNum As Single
randomNum = Rnd
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_msgbox() {
        let source = r#"
MsgBox "Random: " & Rnd
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_debug_print() {
        let source = r#"
Debug.Print Rnd()
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_select_case() {
        let source = r#"
Dim value As Single
value = Rnd
Select Case value
    Case Is < 0.33
        result = "Low"
    Case Is < 0.67
        result = "Medium"
    Case Else
        result = "High"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_class_usage() {
        let source = r#"
Private m_randomValue As Single

Public Sub GenerateRandom()
    m_randomValue = Rnd
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_with_statement() {
        let source = r#"
With dataObject
    .RandomValue = Rnd
End With
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_elseif() {
        let source = r#"
Dim r As Single
r = Rnd
If r < 0.25 Then
    category = 1
ElseIf r < 0.5 Then
    category = 2
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_for_loop() {
        let source = r#"
For i = 1 To 10
    randomNumbers(i) = Rnd
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_do_while() {
        let source = r#"
Do While Rnd < 0.95
    count = count + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_do_until() {
        let source = r#"
Do Until Rnd > 0.9
    attempts = attempts + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_while_wend() {
        let source = r#"
While Rnd < 0.8
    iterations = iterations + 1
Wend
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_parentheses() {
        let source = r#"
Dim val As Single
val = (Rnd)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_iif() {
        let source = r#"
Dim result As String
result = IIf(Rnd > 0.5, "Win", "Lose")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_array_assignment() {
        let source = r#"
Dim values(10) As Single
values(i) = Rnd
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_property_assignment() {
        let source = r#"
Set obj = New RandomData
obj.Value = Rnd
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_function_argument() {
        let source = r#"
Call ProcessValue(Rnd)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_concatenation() {
        let source = r#"
Dim msg As String
msg = "Value: " & Rnd
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_with_int() {
        let source = r#"
Dim diceRoll As Integer
diceRoll = Int(Rnd * 6) + 1
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Int"));
    }

    #[test]
    fn test_rnd_comparison() {
        let source = r#"
If Rnd < 0.3 Then
    status = "Rare"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_multiple_calls() {
        let source = r#"
Dim x As Single, y As Single
x = Rnd
y = Rnd
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_error_handling() {
        let source = r#"
On Error Resume Next
Dim randomVal As Single
randomVal = Rnd
If Err.Number <> 0 Then
    randomVal = 0.5
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_rnd_on_error_goto() {
        let source = r#"
Sub GenerateRandom()
    On Error GoTo ErrorHandler
    Dim r As Single
    r = Rnd
    Exit Sub
ErrorHandler:
    MsgBox "Error generating random number"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rnd"));
        assert!(text.contains("Identifier"));
    }
}
