using ExcelTest1.DTOs;
using ExcelTest1.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelTest1.Controllers;

[Route("api/[controller]")]
[ApiController]
public class AuthController : ControllerBase
{
    private readonly AuthService _authService;

    public AuthController(AuthService authService)
    {
        _authService = authService;
    }

    [HttpPost("[action]")]
    public async Task<IActionResult> Login([FromBody] LoginDto loginDto)
    {
        try
        {
            var result = await _authService.LoginAsync(loginDto);

            if (result == null)
            {
                return Unauthorized(new { message = "Username yoki password noto'g'ri" });
            }

            return Ok(result);
        }
        catch (Exception ex)
        {
            return BadRequest(new { message = ex.Message });
        }
    }

    [HttpPost("[action]")]
    public async Task<IActionResult> Register([FromBody] RegisterDto registerDto)
    {
        try
        {
            var result = await _authService.RegisterAsync(registerDto);

            if (result == null)
            {
                return BadRequest(new { message = "Ro'yxatdan o'tishda xatolik" });
            }

            return Ok(result);
        }
        catch (InvalidOperationException ex)
        {
            return BadRequest(new { message = ex.Message });
        }
        catch (Exception ex)
        {
            return BadRequest(new { message = "Ro'yxatdan o'tishda xatolik yuz berdi" });
        }
    }

    [HttpPost("[action]")]
    public async Task<IActionResult> RegisterAdmin([FromBody] AdminRegisterDto adminRegisterDto)
    {
        try
        {
            // AdminRegisterDto'ni RegisterDto'ga convert qilish
            var registerDto = new RegisterDto
            {
                Username = adminRegisterDto.Username,
                Email = adminRegisterDto.Email,
                Password = adminRegisterDto.Password,
                ConfirmPassword = adminRegisterDto.ConfirmPassword,
                Role = "Admin" // Admin rolini o'rnatish
            };

            var result = await _authService.RegisterAsync(registerDto);
            if (result == null)
            {
                return BadRequest(new { message = "Admin yaratishda xatolik" });
            }
            return Ok(result);
        }
        catch (InvalidOperationException ex)
        {
            return BadRequest(new { message = ex.Message });
        }
        catch (Exception ex)
        {
            return BadRequest(new { message = "Admin yaratishda xatolik yuz berdi" });
        }
    }
}
