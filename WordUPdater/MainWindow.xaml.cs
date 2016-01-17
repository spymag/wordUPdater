<Window x:Class="WordUPdater.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:i="http://schemas.microsoft.com/expression/2010/interactivity"
        xmlns:implementation="clr-namespace:MyAttachedBehaviors"
        Title="Word Updater Beta V1.0.0.0"
        MinHeight="360"
        Height="300"
        MinWidth="484"
        SizeToContent="Width"
        Icon="Resources/AppIcon.ico" 
        DataContext="{Binding Main,Source={StaticResource Locator}}">

    <Window.Resources>
        <ResourceDictionary Source="MainWindowResources.xaml"/>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width=".1*" />
            <ColumnDefinition Width="1.3*" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" MinHeight="50" />
            <RowDefinition Height="14*" MinHeight="80" />
        </Grid.RowDefinitions>

        <Menu Style="{StaticResource MenuStyle}">
            <MenuItem Header="_File ">
                <MenuItem Header="_Exit" Name="mnuExit" ToolTip="Close the application." Command="{Binding ExitApplication}"/>
            </MenuItem>
            <MenuItem Header="_Help ">
                <MenuItem Header="_About " ToolTip="Shows Help Message."  Command="{Binding ShowHelp}"/>
            </MenuItem>
        </Menu>

        <Button Command="{Binding SelectImagesFolder}" Grid.Row="1" Content="SelectImagesFolder" Style="{StaticResource ButtonStyle}">
            <Button.ToolTip>
                <StackPanel Margin="5">
                    <TextBlock Style="{StaticResource TextBlockToolTipStyle}">Select the folder with the images. <LineBreak/> The naming convention should look like below:</TextBlock>
                    <Image Source="Resources/ImagesFolder.bmp"/>
                </StackPanel>
            </Button.ToolTip>
        </Button>

        <Button Command="{Binding SelectWordFile}" Content="SelectWordFile" Grid.Row="2" Style="{StaticResource ButtonStyle}">
            <Button.ToolTip>
                <StackPanel Margin="5">
                    <TextBlock Style="{StaticResource TextBlockToolTipStyle}">Select the word (docx) file.</TextBlock>
                    <Image Source="Resources/SelectWordFile.bmp"/>
                </StackPanel>
            </Button.ToolTip>
        </Button>

        <Button Command="{Binding UpdatePictures}" Content="UpdatePictures" Grid.Row="3" ToolTip="Sync pictures between folder and docx file" Style="{StaticResource ButtonStyle}"/>
        <Button Command="{Binding SaveBitmapFromClipboard}" Content="Save Bitmap from Clipboard" Grid.Row="4" Style="{StaticResource ButtonStyle}">
            <Button.ToolTip>
                <StackPanel Margin="5">
                    <TextBlock Style="{StaticResource TextBlockToolTipStyle}">Create a bmp image file of your current clipboard content.<LineBreak/> The image will be saved to the selected image folder as shown below.</TextBlock>
                    <Image Source="Resources/SaveImageFromClipboard.bmp"/>
                </StackPanel>

            </Button.ToolTip>


        </Button>
        <Label Content="_Custom File Number:" Grid.Row="5" Style="{StaticResource LabelStyle}" Target="{Binding ElementName=FileToReplaceInput}" Grid.ColumnSpan="1" Grid.Column="1" Width="121">
            <Label.ToolTip>
                <StackPanel Margin="5">
                    <TextBlock>It will create a file with the format: "ImageXXX" where XXX is the entered value.<LineBreak/> Leave it blank for default, where default is asceding numbering based on higest value.</TextBlock>
                </StackPanel>
            </Label.ToolTip>


        </Label>

        <TextBox Text="{Binding FileToReplace}" Grid.Column="3" Grid.Row="5" Style="{StaticResource TextBoxReadOnlyStyle}" IsReadOnly="False" Name="FileToReplaceInput"/>
        <TextBox Text="{Binding ImagesPath}" Grid.Row="1" Grid.Column="3" Name="Textbox1" Style="{StaticResource TextBoxReadOnlyStyle}"/>
        <TextBox Text="{Binding DocxPath}" Grid.Row="2" Grid.Column="3" Style="{StaticResource TextBoxReadOnlyStyle}"/>
        <TextBox Text="{Binding PicFromClipboardPath}" Grid.Row="3" Grid.Column="3" Style="{StaticResource TextBoxReadOnlyStyle}"/>
        <TextBox Text="{Binding BackupFolder}" Grid.Row="7" Grid.Column="3" Name="BackupFolderTextBox" IsReadOnly="False" Style="{StaticResource TextBoxReadOnlyStyle}"/>

        <ScrollViewer Grid.Row="8" Grid.ColumnSpan="4" Margin="2">
            <i:Interaction.Behaviors>
                <implementation:AutoScrollBehavior />
            </i:Interaction.Behaviors>
            <TextBox Text="{Binding SystemMessage}" Height="auto" Style="{StaticResource TextBoxReadOnlyStyle}"/>
        </ScrollViewer>





        <Label Content="_BackUp Path:" Target="{Binding ElementName=BackupFolderTextBox}" Grid.Row="7" Grid.Column="1" ToolTip="Will Appear upon running updating a word file." Style="{StaticResource LabelStyle}" Grid.ColumnSpan="1" Width="79"/>
    </Grid>
</Window>
