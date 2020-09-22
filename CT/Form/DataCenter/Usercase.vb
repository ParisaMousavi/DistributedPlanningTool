

Namespace Form.DataCenter

    Friend NotInheritable Class Usercase

        Public ReadOnly Property UsercaseAllocatedSequence As Long

            Get
                Try

                    UsercaseAllocatedSequence = 0
                Catch

                    UsercaseAllocatedSequence = 0

                End Try

            End Get

        End Property


        Public ReadOnly Property UsercaseStartDate As Date

            Get
                Try

                    UsercaseStartDate = Date.Now

                Catch

                    UsercaseStartDate = Nothing

                End Try

            End Get

        End Property



    End Class

End Namespace
