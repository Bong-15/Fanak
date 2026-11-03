using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Fanak
{
    // Shared look for every window. It is parsed at runtime because this build has no XAML compiler.
    internal static class Theme
    {
        private const string Xaml = @"
<ResourceDictionary xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
                    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
  <SolidColorBrush x:Key='Accent' Color='#B9520A'/>
  <SolidColorBrush x:Key='AccentDark' Color='#99430A'/>
  <SolidColorBrush x:Key='Line' Color='#D3D3DA'/>
  <SolidColorBrush x:Key='Muted' Color='#5C5C66'/>
  <SolidColorBrush x:Key='Surface' Color='#F6F6F8'/>
  <SolidColorBrush x:Key='Ink' Color='#1F1F24'/>

  <Style TargetType='Button'>
    <Setter Property='Background' Value='White'/>
    <Setter Property='Foreground' Value='{StaticResource Ink}'/>
    <Setter Property='BorderBrush' Value='{StaticResource Line}'/>
    <Setter Property='Padding' Value='14,6'/>
    <Setter Property='MinWidth' Value='72'/>
    <Setter Property='Cursor' Value='Hand'/>
    <Setter Property='FocusVisualStyle' Value='{x:Null}'/>
    <Setter Property='Template'>
      <Setter.Value>
        <ControlTemplate TargetType='Button'>
          <Border CornerRadius='6' BorderThickness='1' Background='{TemplateBinding Background}'
                  BorderBrush='{TemplateBinding BorderBrush}' Padding='{TemplateBinding Padding}'>
            <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center' RecognizesAccessKey='False'/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property='IsMouseOver' Value='True'><Setter Property='Background' Value='#EFEFF2'/></Trigger>
      <Trigger Property='IsPressed' Value='True'><Setter Property='Background' Value='#E3E3E8'/></Trigger>
      <Trigger Property='IsKeyboardFocused' Value='True'><Setter Property='BorderBrush' Value='{StaticResource Accent}'/></Trigger>
      <Trigger Property='IsEnabled' Value='False'><Setter Property='Opacity' Value='0.5'/></Trigger>
    </Style.Triggers>
  </Style>

  <Style x:Key='PrimaryButton' TargetType='Button' BasedOn='{StaticResource {x:Type Button}}'>
    <Setter Property='Background' Value='{StaticResource Accent}'/>
    <Setter Property='Foreground' Value='White'/>
    <Setter Property='BorderBrush' Value='{StaticResource Accent}'/>
    <Setter Property='FontWeight' Value='SemiBold'/>
    <Style.Triggers>
      <Trigger Property='IsMouseOver' Value='True'><Setter Property='Background' Value='{StaticResource AccentDark}'/></Trigger>
      <Trigger Property='IsPressed' Value='True'><Setter Property='Background' Value='{StaticResource AccentDark}'/></Trigger>
      <Trigger Property='IsKeyboardFocused' Value='True'><Setter Property='BorderBrush' Value='{StaticResource Ink}'/></Trigger>
    </Style.Triggers>
  </Style>

  <Style TargetType='TabControl'>
    <Setter Property='BorderThickness' Value='0,1,0,0'/>
    <Setter Property='BorderBrush' Value='{StaticResource Line}'/>
    <Setter Property='Padding' Value='0'/>
    <Setter Property='Background' Value='White'/>
  </Style>

  <Style TargetType='TabItem'>
    <Setter Property='Foreground' Value='{StaticResource Muted}'/>
    <Setter Property='FocusVisualStyle' Value='{x:Null}'/>
    <Setter Property='Template'>
      <Setter.Value>
        <ControlTemplate TargetType='TabItem'>
          <Border x:Name='Tab' BorderThickness='0,0,0,3' BorderBrush='Transparent' Background='Transparent' Padding='16,9'>
            <ContentPresenter ContentSource='Header' RecognizesAccessKey='False'/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property='IsSelected' Value='True'>
              <Setter TargetName='Tab' Property='BorderBrush' Value='{StaticResource Accent}'/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property='IsSelected' Value='True'>
        <Setter Property='Foreground' Value='{StaticResource Ink}'/>
        <Setter Property='FontWeight' Value='SemiBold'/>
      </Trigger>
      <Trigger Property='IsMouseOver' Value='True'><Setter Property='Foreground' Value='{StaticResource Ink}'/></Trigger>
    </Style.Triggers>
  </Style>

  <Style x:Key='Card' TargetType='Border'>
    <Setter Property='Background' Value='{StaticResource Surface}'/>
    <Setter Property='CornerRadius' Value='8'/>
    <Setter Property='Padding' Value='16,14'/>
    <Setter Property='Margin' Value='0,0,0,12'/>
  </Style>

  <Style x:Key='KeyChip' TargetType='Border'>
    <Setter Property='Background' Value='White'/>
    <Setter Property='BorderBrush' Value='{StaticResource Line}'/>
    <Setter Property='BorderThickness' Value='1,1,1,3'/>
    <Setter Property='CornerRadius' Value='6'/>
    <Setter Property='Padding' Value='14,5'/>
    <Setter Property='MinWidth' Value='56'/>
  </Style>

  <Style x:Key='Heading' TargetType='TextBlock'>
    <Setter Property='FontSize' Value='14'/>
    <Setter Property='FontWeight' Value='SemiBold'/>
    <Setter Property='Foreground' Value='{StaticResource Ink}'/>
  </Style>

  <Style x:Key='Hint' TargetType='TextBlock'>
    <Setter Property='FontSize' Value='12'/>
    <Setter Property='Foreground' Value='{StaticResource Muted}'/>
    <Setter Property='TextWrapping' Value='Wrap'/>
    <Setter Property='LineHeight' Value='18'/>
  </Style>
</ResourceDictionary>";

        private static ImageSource _icon;

        public static void Apply(Application app)
        {
            app.Resources.MergedDictionaries.Add((ResourceDictionary)XamlReader.Parse(Xaml));
        }

        public static Stream OpenIconStream()
        {
            return Assembly.GetExecutingAssembly().GetManifestResourceStream("Fanak.ico");
        }

        public static ImageSource AppIcon
        {
            get
            {
                if (_icon == null)
                {
                    using (Stream stream = OpenIconStream())
                    {
                        _icon = BitmapFrame.Create(stream, BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
                    }
                }
                return _icon;
            }
        }
    }
}
